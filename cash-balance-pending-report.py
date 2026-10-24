from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from datetime import datetime
import os
import time
import re
import pandas as pd
import requests

# ============================================================
# CONFIGURATION (BACK TO SECRETS - NO HARDCODING)
# ============================================================
# 💡 यह आपके GitHub Secrets से डेटा उठाएगा, आपको यहाँ कुछ नहीं बदलना है
USERNAME = os.getenv("ACEN_USERNAME")
PASSWORD = os.getenv("ACEN_PASSWORD")
LOGIN_URL = os.getenv("ACEN_LOGIN_URL")

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

BASE_DIR = os.getcwd()
OUTPUT_FILE_PATH = os.path.join(BASE_DIR, "Cash_Balance_Pending_Report.xlsx")

# ============================================================
# TELEGRAM FUNCTION
# ============================================================
def send_telegram_report(message, excel_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram credentials missing.")
        return
    base_url = f"https://telegram.org{TELEGRAM_BOT_TOKEN}"
    if excel_path and os.path.exists(excel_path):
        try:
            requests.post(
                f"{base_url}/sendDocument",
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": message, "parse_mode": "Markdown"},
                files={"document": open(excel_path, "rb")},
                timeout=60
            )
            print("Telegram पर रिपोर्ट भेज दी गई है।")
        except Exception as e:
            print("Telegram Error:", e)

# ============================================================
# MAIN AUTOMATION
# ============================================================
def run_cash_balance_report_job():
    print("जॉब शुरू हो रही है...")
    raw_download_path = os.path.join(BASE_DIR, "temp_raw_data.xlsx")

    # 💡 1 सेकंड में बंद होने से रोकने के लिए डायरेक्ट चेक हटा दिया गया है
    with sync_playwright() as p:
        print("क्रोमियम लॉन्च किया जा रहा है...")
        browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = context.new_page()
        
        print(f"लॉगिन पेज ओपन किया जा रहा है: {LOGIN_URL}")
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
        time.sleep(5)
        
        # वेबसाइट iframes का उपयोग करती है, इसलिए सही फ़्रेम सेलेक्ट करें
        target = page
        if len(page.frames) > 1:
            for f in page.frames:
                if f.name or f.url:
                    target = f
                    break

        # लॉगिन प्रोसेस
        try:
            print("लॉगिन डिटेल्स भरी जा रही हैं...")
            target.locator("input[type='text'], input[type='email'], input[name='username'], #username").first.fill(str(USERNAME))
            target.locator("input[type='password'], #password").first.fill(str(PASSWORD))
            time.sleep(1)
            target.locator("button[type='submit'], input[type='submit'], button:has-text('Login')").first.click()
            print("लॉगिन सबमिट हो गया।")
            time.sleep(8)
        except Exception as e:
            print("लॉगिन स्टेप स्किप या फेल:", e)

        # लॉगिन के बाद मुख्य फ़्रेम को फिर से टारगेट करें
        for f in page.frames:
            if "main" in f.name.lower() or "report" in f.name.lower() or len(f.name) > 0:
                target = f
                break

        # Reports बटन पर क्लिक
        print("Reports बटन पर जाया जा रहा है...")
        target.get_by_text("Reports", exact=True).wait_for(state="visible", timeout=60000)
        target.get_by_text("Reports", exact=True).click()
        time.sleep(4)

        # फ़ाइल डाउनलोड
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
    # DATA PROCESSING & FILTERING (A to Z)
    # ============================================================
    if not os.path.exists(raw_download_path):
        print("Error: एक्सेल फ़ाइल नहीं मिली!")
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
    run_cash_balance_report_job()
