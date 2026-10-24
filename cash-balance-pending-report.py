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

def send_telegram_report(message, excel_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("❌ टेलीग्राम सीक्रेट्स (Secrets) गायब हैं!")
        return
    base_url = f"https://telegram.org{TELEGRAM_BOT_TOKEN}"
    if excel_path and os.path.exists(excel_path):
        print("➡️ टेलीग्राम पर फ़ाइल भेजी जा रही है...")
        try:
            res = requests.post(
                f"{base_url}/sendDocument",
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": message, "parse_mode": "Markdown"},
                files={"document": open(excel_path, "rb")},
                timeout=60
            )
            print(f"टेलीग्राम रिस्पॉन्स: {res.status_code} - {res.text}")
        except Exception as e:
            print("❌ टेलीग्राम भेजने में एरर आया:", e)

def run_cash_balance_report_job():
    print(f"START: जॉब शुरू हुई। लॉगिन URL: {LOGIN_URL}")
    if not USERNAME or not PASSWORD or not LOGIN_URL:
        print("❌ Error: GitHub Secrets में लॉगिन डिटेल्स (Username/Password/URL) गायब हैं!")
        return

    raw_download_path = os.path.join(BASE_DIR, "temp_raw_data.xlsx")

    with sync_playwright() as p:
        print("🔄 क्रोम ब्राउज़र लॉन्च किया जा रहा है...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        try:
            print("🌐 लॉगिन पेज पर जाया जा रहा है...")
            page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
            print("✅ लॉगिन पेज लोड हो गया।")
            
            # 💡 स्क्रीनशॉट 1: लॉगिन पेज कैसा दिख रहा है
            page.screenshot(path="step1_login_page.png")
            
            # --- ऑटोमेशन स्टेप्स ---
            # अगर आपके पोर्टल पर डायरेक्ट रिपोर्ट्स पेज खुलता है या जो भी हो, 
            # यह कोड 'Reports' बटन को ढूँढेगा:
            print("🔍 'Reports' बटन को स्क्रीन पर ढूँढा जा रहा है...")
            reports_btn = page.get_by_text("Reports", exact=True)
            reports_btn.wait_for(state="visible", timeout=30000)
            print("🖱️ 'Reports' बटन मिल गया, इस पर क्लिक किया जा रहा है...")
            reports_btn.click()
            
            # 💡 स्क्रीनशॉट 2: रिपोर्ट्स पर क्लिक करने के बाद क्या खुला
            page.screenshot(path="step2_after_reports.png")
            
            print("📥 एक्सेल डाउनलोड बटन का इंतज़ार किया जा रहा है...")
            # यहाँ कोड मानकर चल रहा है कि डाउनलोड बटन पर 'Download' या 'Export' लिखा होगा
            download_btn = page.locator("button, a").filter(has_text=re.compile(r"download|export", re.IGNORECASE)).first
            
            with page.expect_download(timeout=60000) as download_info:
                download_btn.click()
            
            download = download_info.value
            download.save_as(raw_download_path)
            print("✅ एक्सेल फ़ाइल सफलतापूर्वक डाउनलोड हो गई!")
            
        except Exception as err:
            print(f"❌ ऑटोमेशन के दौरान गड़बड़ हुई: {err}")
            # 💡 फेल होने से ठीक पहले का लाइव स्क्रीनशॉट
            page.screenshot(path="crash_error_view.png")
            print("📸 क्रैश होने से ठीक पहले का स्क्रीनशॉट 'crash_error_view.png' सेव कर लिया गया है।")
            browser.close()
            
            # टेलीग्राम पर एरर और स्क्रीनशॉट भेजें ताकि मोबाइल पर ही पता चल जाए
            send_telegram_report(f"❌ *जॉब फेल हो गई!*\n\nएरर: `{err}`", "crash_error_view.png")
            return
            
        browser.close()

    # ============================================================
    # DATA PROCESSING (PANDAS)
    # ============================================================
    if not os.path.exists(raw_download_path):
        print("❌ Error: डाउनलोड फ़ोल्डर में एक्सेल फ़ाइल नहीं मिली!")
        return

    print("📊 एक्सेल डेटा को रीड और फ़िल्टर किया जा रहा है...")
    df = pd.read_excel(raw_download_path)
    df.columns = df.columns.str.strip()
    
    # कॉलम खोजना
    pending_col = next((c for c in df.columns if 'pending' in c.lower() or 'cash' in c.lower()), None)
    emp_col = next((c for c in df.columns if 'emp' in c.lower() or 'name' in c.lower()), None)
    
    if not pending_col or not emp_col:
        print(f"❌ सही कॉलम नहीं मिले! आपकी शीट में ये कॉलम्स हैं: {df.columns.tolist()}")
        return

    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[emp_col] = df[emp_col].astype(str).str.strip()
    
    # फ़िल्टर: पेंडिंग > 0 (Active + Inactive दोनों)
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
        "📢 *Note:* Active/Inactive दोनों शामिल हैं।"
    )
    
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    run_cash_balance_report_job()
