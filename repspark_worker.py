import os
import time
import traceback
from pathlib import Path

# Evita conflitos de event loop em runners
import nest_asyncio
nest_asyncio.apply()

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


def _log(msg: str):
    print(f"[RS] {msg}", flush=True)


def wait_and_click_xpath_anywhere(page, xpaths, timeout=15000, debug_prefix="debug"):
    """
    Tenta cada XPath no documento principal e em todos os iframes.
    Faz wait_for_selector(state='visible') antes de clicar.
    Se falhar, salva screenshot e retorna False.
    """
    # principal
    for xp in xpaths:
        sel = f"xpath={xp}"
        try:
            page.wait_for_selector(sel, state="visible", timeout=timeout)
            el = page.locator(sel)
            el.scroll_into_view_if_needed(timeout=3000)
            el.click(timeout=timeout)
            return True
        except Exception:
            pass

    # iframes
    for fr in page.frames:
        for xp in xpaths:
            sel = f"xpath={xp}"
            try:
                fr.wait_for_selector(sel, state="visible", timeout=timeout)
                el = fr.locator(sel)
                el.scroll_into_view_if_needed(timeout=3000)
                el.click(timeout=timeout)
                return True
            except Exception:
                pass

    # fallback por texto visível (se existir)
    try:
        page.get_by_text("Export ATS to Excel", exact=False).first.click(timeout=3000)
        return True
    except Exception:
        pass

    # debug screenshot
    ts = int(time.time())
    debug_path = Path(f"{debug_prefix}_{ts}.png")
    try:
        page.screenshot(path=str(debug_path), full_page=True)
        _log(f"DEBUG screenshot salvo: {debug_path}")
    except Exception:
        pass
    return False


def run():
    try:
        download_dir = Path("downloads")
        download_dir.mkdir(exist_ok=True)

        # -------- ENV / Secrets --------
        repspark_url = os.environ.get("REPSPARK_URL", "https://app.repspark.com/_511")
        email        = os.environ["REPSPARK_EMAIL"]
        password     = os.environ["REPSPARK_PASSWORD"]
        gsheet_id    = os.environ["GSHEET_ID"]
        gsheet_tab   = os.environ.get("GSHEET_TAB", "BASE")
        sa_json      = os.environ["GCP_SA_JSON"]   # JSON completo da service account

        products_xpath = os.environ.get(
            "PRODUCTS_XPATH",
            "/html/body/div[2]/div[1]/div/div[1]/div[2]/div[2]/div/ul/li[1]/a"
        )
        export_id_xp = os.environ.get(
            "EXPORT_BTN_ID_XPATH",
            "//*[@id='ctl00_ctl00_cphB_filterMenu_btnExportToExcelFull']"
        )
        export_fb_xp = os.environ.get(
            "EXPORT_FALLBACK_XPATH",
            "/html/body/form/div[3]/div[1]/div[3]/div[1]/div/nav/div/div[2]/div/a/span"
        )
        export_xpaths = [export_id_xp, export_fb_xp]

        # -------- Service Account JSON -> arquivo local --------
        sa_path = Path("sa.json")
        sa_path.write_text(sa_json)

        _log("Iniciando Playwright…")
        # Usar start/stop manual para evitar "Cannot close a running event loop"
        p = sync_playwright().start()
        browser = None
        try:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
            context = browser.new_context(accept_downloads=True, viewport={"width": 1400, "height": 900})
            page = context.new_page()

            _log(f"Acessando: {repspark_url}")
            page.goto(repspark_url, wait_until="domcontentloaded", timeout=120_000)
            page.wait_for_load_state("networkidle", timeout=60_000)

            # ---- Login se necessário
            needs_login = False
            try:
                needs_login = page.get_by_placeholder("Email").count() > 0
            except Exception:
                if "login" in page.url.lower():
                    needs_login = True

            if needs_login:
                _log("Fazendo login…")

                def fill(sel, val):
                    try:
                        page.locator(sel).fill(val, timeout=10_000)
                        return True
                    except PWTimeout:
                        return False

                assert (fill('input[name="email"]', email)
                        or fill('input[type="email"]', email)
                        or fill('input[placeholder*="Email" i]', email)), "Campo de e-mail não encontrado"

                assert (fill('input[name="password"]', password)
                        or fill('input[type="password"]', password)
                        or fill('input[placeholder*="Password" i]', password)), "Campo de senha não encontrado"

                try:
                    page.get_by_role("button", name="Sign in").click(timeout=10_000)
                except Exception:
                    page.keyboard.press("Enter")

                page.wait_for_load_state("networkidle", timeout=60_000)
                _log("Login OK.")

            # ---- Ir para Products
            _log("Abrindo Products…")
            ok_products = wait_and_click_xpath_anywhere(page, [products_xpath], timeout=20000, debug_prefix="debug_products")
            assert ok_products, "Não consegui clicar no menu Products pelos XPaths fornecidos."
            page.wait_for_load_state("networkidle", timeout=60_000)
            page.wait_for_timeout(800)

            # ---- Exportar Excel (com retries e debug)
            _log("Exportando Excel…")
            success = False
            for attempt in range(1, 4):
                try:
                    with page.expect_download(timeout=180_000) as dlinfo:
                        if not wait_and_click_xpath_anywhere(page, export_xpaths, timeout=20000, debug_prefix=f"debug_export_try{attempt}"):
                            raise RuntimeError("Botão de exportação não encontrado pelos XPaths informados.")
                    download = dlinfo.value
                    filename = download.suggested_filename or f"Availability_{int(time.time())}.xlsx"
                    xlsx_path = download_dir / filename
                    download.save_as(str(xlsx_path))
                    _log(f"Download OK: {xlsx_path}")
                    success = True
                    break
                except Exception as e:
                    _log(f"Tentativa {attempt} falhou: {e}")
                    page.wait_for_timeout(1500)

            assert success, "Falha ao acionar o download após múltiplas tentativas."

            # ---- Atualizar Google Sheets
            _log("Atualizando Google Sheets…")
            creds = Credentials.from_service_account_file(
                str(sa_path),
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(gsheet_id)

            try:
                ws = sh.worksheet(gsheet_tab)
            except gspread.WorksheetNotFound:
                ws = sh.add_worksheet(title=gsheet_tab, rows="100", cols="26")

            df = pd.read_excel(xlsx_path, engine="openpyxl")
            values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()

            ws.clear()
            ws.resize(rows=max(len(values), 100), cols=max(len(values[0]), 26))
            ws.update("A1", values, value_input_option="RAW")
            _log("Planilha atualizada com sucesso.")

        finally:
            # Fechamento seguro
            try:
                if browser is not None:
                    browser.close()
            finally:
                p.stop()

    except Exception:
        _log("[ERRO]\n" + traceback.format_exc())
        raise


if __name__ == "__main__":
    run()
