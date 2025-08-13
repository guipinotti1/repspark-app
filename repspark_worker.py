import os
import time
import traceback
from pathlib import Path

# Evita erro "This event loop is already running" em runners/Actions
import nest_asyncio
nest_asyncio.apply()

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


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
        sa_json      = os.environ["GCP_SA_JSON"]   # cole o JSON inteiro no secret

        # XPaths (vêm dos defaults ou de env se você quiser trocar depois)
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

        print("[RS] Iniciando Playwright…", flush=True)
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            print("[RS] Acessando:", repspark_url, flush=True)
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
                print("[RS] Fazendo login…", flush=True)

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
                print("[RS] Login OK.", flush=True)

            # ---- Ir para Products
            print("[RS] Abrindo Products…", flush=True)
            page.locator(f"xpath={products_xpath}").click(timeout=20_000)
            page.wait_for_load_state("networkidle", timeout=60_000)

            # ---- Exportar Excel
            print("[RS] Exportando Excel…", flush=True)
            with page.expect_download(timeout=180_000) as dlinfo:
                clicked = False
                for xp in export_xpaths:
                    try:
                        page.locator(f"xpath={xp}").click(timeout=8_000)
                        clicked = True
                        break
                    except Exception:
                        pass
                assert clicked, "Botão de exportação não encontrado pelos XPaths informados."

            download = dlinfo.value
            filename = download.suggested_filename or f"Availability_{int(time.time())}.xlsx"
            xlsx_path = download_dir / filename
            download.save_as(str(xlsx_path))
            print(f"[RS] Download OK: {xlsx_path}", flush=True)

            context.close()
            browser.close()

        # -------- Atualizar Google Sheets
        print("[RS] Atualizando Google Sheets…", flush=True)
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
        print("[RS] Planilha atualizada com sucesso.", flush=True)

    except Exception:
        print("[RS][ERRO]\n" + traceback.format_exc(), flush=True)
        # Re-raise para o Actions marcar como falha
        raise


if __name__ == "__main__":
    run()
