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
# TELEGRAM FUNCTION (FIXED)
# ============================================================
def send_telegram_report(message, excel_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram credentials missing. Dispatch skipped.")
        return
    
    # बिल्कुल सही API URL
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
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        # 1. पोर्टल लॉगिन प्रोसेस
        print("लॉगिन पेज ओपन किया जा रहा है...")
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
        
        # 💡 ऑटो-लॉगिन: यह कोड इनपुट्स को खुद ढूँढकर यूजरनेम और पासवर्ड भर देगा
        try:
            page.locator("input[type='text'], input[type='email'], input[name='username']").first.fill(USERNAME)
            page.locator("input[type='password']").first.fill(PASSWORD)
            page.locator("button[type='submit'], input[type='submit']").first.click()
            print("लॉगिन क्रेडेंशियल्स सबमिट कर दिए गए हैं।")
            time.sleep(5)
        except Exception as login_err:
            print(f"ऑटो-लॉगिन में दिक्कत आई (शायद पहले से लॉगिन था): {login_err}")
        
        # 2. Reports बटन पर जाना
        print("Reports सेक्शन में जाया जा रहा... बटन ढूँढा जा रहा है...")
        page.get_by_text("Reports", exact=True).wait_for(state="visible", timeout=60000)
        page.get_by_text("Reports", exact=True).click()
        
        # 3. फ़ाइल डाउनलोड प्रोसेस (स्मार्ट डाउनलोड हैंडलर)
        print("डेटा डाउनलोड ट्रिगर किया जा रहा है...")
        try:
            with page.expect_download(timeout=60000) as download_info:
                # यह स्क्रीन पर मौजूद डाउनलोड/एक्सपोर्ट बटन को खुद ढूंढकर क्लिक करेगा
                page.locator("button, a").filter(has_text=re.compile(r"download|export|एक्सेल|excel", re.IGNORECASE)).first.click()
            
            download = download_info.value
            download.save_as(raw_download_path)
            print("डेटा सफलतापूर्वक डाउनलोड हो गया।")
        except Exception as dl_err:
            print(f"डाउनलोड बटन क्लिक करने में एरर: {dl_err}")
            # फॉलबैक: अगर बटन न मिले तो पुराने स्टाइल से आगे बढ़े
            pass
            
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
    
    # कॉलम नाम ऑटो-डिटेक्ट करना (ताकि इनएक्टिव वाले भी मिस न हों)
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

    # टेलीग्राम समरी मेसेज तैयार करना
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
    
    # टेलीग्राम पर सेंड करना (हर हाल में सेंड करेगा)
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    try:
        run_cash_balance_report_job()
    except Exception as e:
        print(f"इस रिपोर्ट जॉब में एरर आया: {e}")
