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
# TELEGRAM FUNCTION (100% FIXED)
# ============================================================
def send_telegram_report(message, excel_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram credentials missing. Dispatch skipped.")
        return
    
    base_url = f"https://telegram.org{TELEGRAM_BOT_TOKEN}"
    
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
    if not USERNAME or not PASSWORD or not LOGIN_URL:
        print("Error: Environment Variables (Secrets) missing!")
        return

    raw_download_path = os.path.join(BASE_DIR, "temp_raw_data.xlsx")

    with sync_playwright() as p:
        print("Browser ओपन किया जा रहा है...")
        # 💡 ब्लॉक होने से बचाने के लिए क्रोमियम आर्गुमेंट्स जोड़े गए हैं
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        
        # 💡 असली कंप्यूटर/मोबाइल जैसा User-Agent सेट किया गया है ताकि वेबसाइट ब्लॉक न करे
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        # 1. पोर्टल लॉगिन प्रोसेस
        print("लॉगिन पेज ओपन किया जा रहा है...")
        page.goto(LOGIN_URL, wait_until="networkidle", timeout=120000)
        
        # 💡 ऑटो-लॉगिन: इनपुट्स को ढूँढकर यूजरनेम और पासवर्ड भरना
        try:
            print("क्रेडेंशियल्स भरे जा रहे हैं...")
            page.locator("input[type='text'], input[type='email'], input[name='username']").first.fill(USERNAME)
            page.locator("input[type='password']").first.fill(PASSWORD)
            time.sleep(1)
            
            # सबमिट बटन पर क्लिक
            submit_btn = page.locator("button[type='submit'], input[type='submit'], button:has-text('Login'), button:has-text('Sign In')").first
            submit_btn.click()
            print("लॉगिन बटन पर सफलतापूर्वक क्लिक किया गया।")
            
            # पेज लोड होने का इंतज़ार
            page.wait_for_load_state("networkidle", timeout=30000)
        except Exception as login_err:
            print(f"लॉगिन स्टेप में रुकावट आई: {login_err}")
        
        # 2. Reports बटन पर जाना
        print("Reports सेक्शन में जाया जा रहा... बटन ढूँढा जा रहा है...")
        try:
            reports_btn = page.get_by_text("Reports", exact=True)
            reports_btn.wait_for(state="visible", timeout=45000)
            reports_btn.click()
            print("Reports बटन पर सफलतापूर्वक क्लिक हो गया।")
            page.wait_for_load_state("networkidle", timeout=20000)
        except Exception as rep_err:
            print(f"Reports बटन ढूँढने में एरर: {rep_err}")
        
        # 3. फ़ाइल डाउनलोड प्रोसेस (मजबूत हैंडलर)
        print("डेटा डाउनलोड ट्रिगर किया जा रहा है...")
        try:
            with page.expect_download(timeout=60000) as download_info:
                # एक्सेल या डाउनलोड बटन को खुद ढूंढकर क्लिक करेगा
                page.locator("button, a").filter(has_text=re.compile(r"download|export|excel|एक्सेल", re.IGNORECASE)).first.click()
            
            download = download_info.value
            download.save_as(raw_download_path)
            print("डेटा सफलतापूर्वक डाउनलोड हो गया।")
        except Exception as dl_err:
            print(f"डाउनलोड प्रोसेस में एरर: {dl_err}")
            
        browser.close()

    # ============================================================
    # PANDAS DATA PROCESSING & FILTERING (A to Z)
    # ============================================================
    if not os.path.exists(raw_download_path):
        print("Error: डाउनलोड की गई फ़ाइल नहीं मिली!")
        return

    print("डाउनलोड की गई फ़ाइल को फ़िल्टर और सॉर्ट किया जा रहा है...")
    df = pd.read_excel(raw_download_path)
    df.columns = df.columns.str.strip()
    
    # कॉलम नाम ऑटो-डिटेक्ट करना
    pending_col = next((c for c in df.columns if 'pending' in c.lower() or 'cash' in c.lower()), 'Pending')
    emp_col = next((c for c in df.columns if 'emp' in c.lower() or 'name' in c.lower()), 'Employee')
    
    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[emp_col] = df[emp_col].astype(str).str.strip()
    
    # फ़िल्टर: सिर्फ 0 से बड़े पेंडिंग वाले (Active + Inactive दोनों शामिल)
    filtered_df = df[df[pending_col] > 0].copy()
    
    # ऑटो सॉर्टिंग: A to Z नाम के क्रम में लगाना
    filtered_df = filtered_df.sort_values(by=emp_col, ascending=True)
    
    # नई एक्सेल फ़ाइल सेव करना
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
        "📢 *Note:* इस लिस्ट में Active और Inactive दोनों कर्मचारी शामिल हैं। 0 पेंडिंग वाले लोगों को हटा दिया गया है और डेटा को नाम के अनुसार A to Z व्यवस्थित किया गया है।"
    )
    
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    run_cash_balance_report_job()
