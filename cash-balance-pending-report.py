from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from datetime import datetime
import os
import time
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import requests

# ============================================================
# CONFIGURATION (CREDENTIALS FROM SECRETS)
# ============================================================
USERNAME = os.getenv("ACEN_USERNAME")
PASSWORD = os.getenv("ACEN_PASSWORD")
LOGIN_URL = os.getenv("ACEN_LOGIN_URL")

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

BASE_DIR = os.getcwd()
# आपकी पसंद के अनुसार फाइनल एक्सेल रिपोर्ट का नाम
OUTPUT_FILE_PATH = os.path.join(BASE_DIR, "Cash_Balance_Pending_Report.xlsx")

# ============================================================
# TELEGRAM FUNCTIONS (YOUR ORIGINAL WORKING CODE WITH FIX)
# ============================================================
def telegram_request(url, data=None, files=None):
    try:
        response = requests.post(url, data=data, files=files, timeout=60)
        if not response.ok:
            print("Telegram API Error:", response.status_code, response.text)
            return False
        return True
    except Exception as e:
        print("Telegram Request Exception:", e)
        return False

def send_telegram_report(bot_token, chat_id, message, excel_path=None):
    if not bot_token or not chat_id:
        print("Telegram credentials missing. Dispatch skipped.")
        return

    # सही टेलीग्राम API लिंक फिक्स कर दिया है
    base_url = f"https://telegram.org{bot_token}"

    # ========================================================
    # SEND EXCEL FILE & CAPTION TOGETHER (JUST LIKE BEFORE)
    # ========================================================
    if excel_path and os.path.exists(excel_path):
        print("Sending Excel file with report caption to Telegram...")
        try:
            with open(excel_path, "rb") as document:
                telegram_request(
                    f"{base_url}/sendDocument",
                    data={"chat_id": chat_id, "caption": message, "parse_mode": "Markdown"},
                    files={"document": document}
                )
        except Exception as e:
            print("Excel Telegram Error:", e)

    print("Telegram dispatch completed.")

# ============================================================
# MAIN AUTOMATION & FILTER PROCESS
# ============================================================
def run_cash_balance_report_job():
    if not USERNAME or not PASSWORD or not LOGIN_URL:
        print("Error: Environment Variables (Secrets) missing!")
        return

    raw_download_path = os.path.join(BASE_DIR, "temp_raw_data.xlsx")

    with sync_playwright() as p:
        print("Browser ओपन किया जा रहा है...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        # 1. पोर्टल लॉगिन प्रोसेस
        print("लॉगिन पेज ओपन किया जा रहा है...")
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
        time.sleep(5)
        
        # सही फ़्रेम टारगेट करना (चूँकि वेबसाइट iframe यूज़ करती है)
        target = page
        if len(page.frames) > 1:
            for f in page.frames:
                if f.name or f.url:
                    target = f
                    break

        try:
            target.locator("input[type='text'], input[type='email'], input[name='username'], #username").first.fill(str(USERNAME))
            target.locator("input[type='password'], #password").first.fill(str(PASSWORD))
            target.locator("button[type='submit'], input[type='submit'], button:has-text('Login')").first.click()
            print("लॉगिन सबमिट हो गया।")
            time.sleep(10) # लॉगिन के बाद पूरा लोड होने का 10 सेकंड का गैप
        except Exception as e:
            print("लॉगिन स्टेप स्किप या फेल:", e)

        for f in page.frames:
            if "main" in f.name.lower() or "report" in f.name.lower() or len(f.name) > 0:
                target = f
                break

        # 2. Reports बटन पर जाना
        print("Reports बटन पर जाया जा रहा है...")
        target.get_by_text("Reports", exact=True).wait_for(state="visible", timeout=60000)
        target.get_by_text("Reports", exact=True).click()
        time.sleep(4)

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
            with page.expect_download(timeout=30000) as download_info:
                page.locator("a:has-text('Download'), button:has-text('Download')").first.click()
            download = download_info.value
            download.save_as(raw_download_path)
            
        browser.close()

    # ============================================================
    # PANDAS DATA PROCESSING & FILTERING (YOUR NEW REQUEST)
    # ============================================================
    if not os.path.exists(raw_download_path):
        print("Error: एक्सेल फ़ाइल नहीं मिली! डाउनलोड फेल हुआ।")
        return

    print("डाउनलोड की गई फ़ाइल को फ़िल्टर और सॉर्ट किया जा रहा है...")
    df = pd.read_excel(raw_download_path)
    df.columns = df.columns.str.strip()
    
    # कॉलम नाम ऑटो-डिटेक्ट करना (ताकि इनएक्टिव कर्मचारी भी मिस न हों)
    pending_col = next((c for c in df.columns if 'pending' in c.lower() or 'cash' in c.lower()), 'Pending')
    emp_col = next((c for c in df.columns if 'emp' in c.lower() or 'name' in c.lower()), 'Employee')
    
    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[emp_col] = df[emp_col].astype(str).str.strip()
    
    # फ़िल्टर: सिर्फ Pending > 0 (Active + Inactive दोनों शामिल रहेंगे)
    filtered_df = df[df[pending_col] > 0].copy()
    
    # ऑटो सॉर्टिंग: कर्मचारियों के नाम के हिसाब से A to Z क्रम में लगाना
    filtered_df = filtered_df.sort_values(by=emp_col, ascending=True)
    
    # नई एक्सेल फ़ाइल सेव करना
    filtered_df.to_excel(OUTPUT_FILE_PATH, index=False)
    print(f"नई फ़िल्टर्ड रिपोर्ट सेव हो गई: {OUTPUT_FILE_PATH}")
    
    if os.path.exists(raw_download_path):
        os.remove(raw_download_path)

    # टेलीग्राम समरी मेसेज
    total_cases = len(filtered_df)
    total_amount = filtered_df[pending_col].sum()
    
    summary_message = (
        "🌅 *Cash Balance Pending Report (A to Z)*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👥 *Total Pending Employees:* {total_cases}\n"
        f"💰 *Total Pending Amount:* ₹{total_amount:,.2f}\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "📢 *Note:* इस लिस्ट में Active और Inactive दोनों कर्मचारी शामिल हैं। 0 पेंडिंग वाले हटा दिए गए हैं और डेटा A to Z नाम के क्रम में है।"
    )
    
    # 4. टेलीग्राम पर फाइनल सेंड करना (मैसेज + एक्सेल दोनों एक साथ)
    send_telegram_report(TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    run_cash_balance_report_job()
