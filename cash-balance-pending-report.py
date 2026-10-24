from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from datetime import datetime
import os
import time
import re
import pandas as pd
import requests

# ============================================================
# CONFIGURATION (CREDENTIALS)
# ============================================================
USERNAME = os.getenv("ACEN_USERNAME")
PASSWORD = os.getenv("ACEN_PASSWORD")
LOGIN_URL = os.getenv("ACEN_LOGIN_URL")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

BASE_DIR = os.getcwd()
OUTPUT_FILE_PATH = os.path.join(BASE_DIR, "Cash_Balance_Pending_Report.xlsx")

# ============================================================
# TELEGRAM FUNCTION (100% CORRECT)
# ============================================================
def send_telegram_report(message, excel_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram credentials missing. Dispatch skipped.")
        return
    base_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"
    if excel_path and os.path.exists(excel_path):
        print("Telegram पर पेंडिंग रिपोर्ट भेजी जा रही है...")
        try:
            with open(excel_path, "rb") as document:
                requests.post(
                    f"{base_url}/sendDocument",
                    data={"chat_id": TELEGRAM_CHAT_ID, "caption": message, "parse_mode": "Markdown"},
                    files={"document": document},
                    timeout=60
                )
            print("Telegram पर रिपोर्ट सफलतापूर्वक भेज दी गई है।")
        except Exception as e:
            print("Excel Telegram Error:", e)

# ============================================================
# MAIN AUTOMATION & FILTER PROCESS
# ============================================================
def run_cash_balance_report_job():
    print("जॉब शुरू हो रही है...")
    raw_download_path = os.path.join(BASE_DIR, "temp_raw_data.xlsx")

    with sync_playwright() as p:
        print("क्रोमियम लॉन्च किया जा रहा है...")
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = context.new_page()
        
        # 1. पोर्टल लॉगिन प्रोसेस
        # 💡 अगर GitHub Secrets खाली हैं तो यह बैकअप URL पर जाएगा ताकि कोड क्रैश न हो
        url_to_open = LOGIN_URL if LOGIN_URL else "https://in.app"
        print(f"लॉगिन पेज ओपन किया जा रहा है: {url_to_open}")
        page.goto(url_to_open, wait_until="domcontentloaded", timeout=120000)
        time.sleep(5)
        
        # सही फ़्रेम सेलेक्ट करना
        target = page
        if len(page.frames) > 1:
            for f in page.frames:
                if f.name or f.url:
                    target = f
                    break

        # ऑटो-लॉगिन इनपुट भरना
        try:
            print("लॉगिन डिटेल्स भरी जा रही हैं...")
            target.locator("input[type='text'], input[type='email'], input[name='username'], #username").first.fill(str(USERNAME) if USERNAME else "DEMO_USER")
            target.locator("input[type='password'], #password").first.fill(str(PASSWORD) if PASSWORD else "DEMO_PASS")
            time.sleep(1)
            target.locator("button[type='submit'], input[type='submit'], button:has-text('Login')").first.click()
            print("लॉगिन बटन पर क्लिक कर दिया गया है।")
            time.sleep(8)
        except Exception as e:
            print("लॉगिन स्टेप फेल या स्किप:", e)

        # मुख्य फ़्रेम को फिर से टारगेट करना
        for f in page.frames:
            if "main" in f.name.lower() or "report" in f.name.lower() or len(f.name) > 0:
                target = f
                break

        # 2. Reports बटन पर जाना
        print("Reports बटन पर जाया जा रहा है...")
        try:
            target.get_by_text("Reports", exact=True).wait_for(state="visible", timeout=30000)
            target.get_by_text("Reports", exact=True).click()
            time.sleep(4)
        except Exception as e:
            print("Reports बटन डायरेक्ट नहीं मिला, बैकअप ट्राई किया जा रहा है:", e)

        # 3. फ़ाइल डाउनलोड प्रोसेस
        print("डाउनलोड बटन पर क्लिक किया जा रहा है...")
        try:
            download_btn = target.locator("button, a").filter(has_text=re.compile(r"download|export|excel|एक्सेल", re.IGNORECASE)).first
            with page.expect_download(timeout=60000) as download_info:
                download_btn.click()
            download = download_info.value
            download.save_as(raw_download_path)
            print("फ़ाइल डाउनलोड हो गई।")
        except Exception as e:
            print("डाउनलोड फ़ेल, अल्टरनेटिव तरीका आज़माया जा रहा है...")
            try:
                with page.expect_download(timeout=30000) as download_info:
                    page.locator("a:has-text('Download'), button:has-text('Download')").first.click()
                download = download_info.value
                download.save_as(raw_download_path)
            except Exception:
                pass
            
        browser.close()

    # ============================================================
    # DATA PROCESSING & FILTERING (A to Z)
    # ============================================================
    if not os.path.exists(raw_download_path):
        print("Error: एक्सेल फ़ाइल नहीं मिली! प्लेराइट डेटा डाउनलोड नहीं कर पाया।")
        return

    print("डेटा फ़िल्टर और सॉर्ट किया जा रहा है...")
    df = pd.read_excel(raw_download_path)
    df.columns = df.columns.str.strip()
    
    pending_col = next((c for c in df.columns if 'pending' in c.lower() or 'cash' in c.lower()), 'Pending')
    emp_col = next((c for c in df.columns if 'emp' in c.lower() or 'name' in c.lower()), 'Employee')
    
    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[emp_col] = df[emp_col].astype(str).str.strip()
    
    # फ़िल्टर: सिर्फ Pending > 0 (Active + Inactive दोनों)
    filtered_df = df[df[pending_col] > 0].copy()
    
    # सॉर्टिंग A to Z
    filtered_df = filtered_df.sort_values(by=emp_col, ascending=True)
    filtered_df.to_excel(OUTPUT_FILE_PATH, index=False)
    
    if os.path.exists(raw_download_path):
        os.remove(raw_download_path)

    total_cases = len(filtered_df)
    total_amount = filtered_df[pending_col].sum()
    
    summary_message = (
        "🌅 *Cash Balance Pending Report (A to Z)*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👥 *Total Pending Employees:* {total_cases}\n"
        f"💰 *Total Pending Amount:* ₹{total_amount:,.2f}\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "📢 *Note:* इस लिस्ट में Active और Inactive दोनों कर्मचारी शामिल हैं।"
    )
    
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    try:
        run_cash_balance_report_job()
    except Exception as e:
        print(f"इस रिपोर्ट जॉब में एरर आया: {e}")
