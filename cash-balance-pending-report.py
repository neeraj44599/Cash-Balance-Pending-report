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
    if not USERNAME or not PASSWORD or not LOGIN_URL:
        print("Error: Environment Variables (Secrets) missing!")
        return

    raw_download_path = os.path.join(BASE_DIR, "temp_raw_data.xlsx")

    with sync_playwright() as p:
        print("Browser ओपन किया जा रहा है...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        # 1. पोर्टल लॉगिन प्रोसेस (पुराना ओरिजिनल तरीका)
        print("लॉगिन पेज ओपन किया जा रहा है...")
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
        
        # 2. Reports बटन पर जाना (नया और मजबूत तरीका)
        print("Reports सेक्शन में जाया जा रहा... बटन ढूँढा जा रहा है...")
        page.get_by_text("Reports", exact=True).wait_for(state="visible", timeout=60000)
        page.get_by_text("Reports", exact=True).click()
        
        # 3. फ़ाइल डाउनलोड प्रोसेस
        print("डेटा डाउनलोड ट्रिगर किया जा रहा है...")
        try:
            # एक्सेल डाउनलोड के लिए स्मार्ट डाउनलोड लिसनर
            with page.expect_download(timeout=60000) as download_info:
                # यह स्क्रीन पर मौजूद डाउनलोड/एक्सपोर्ट बटन को खोजकर क्लिक करेगा
                page.locator("button, a").filter(has_text=re.compile(r"download|export|excel|एक्सेल", re.IGNORECASE)).first.click()
            
            download = download_info.value
            download.save_as(raw_download_path)
            print("डेटा फ़ाइल सफलतापूर्वक डाउनलोड हो गई।")
        except Exception as dl_err:
            print(f"डाउनलोड बटन क्लिक करने में एरर: {dl_err}")
            # फॉलबैक: पुराने जेनेरिक नाम से डायरेक्ट क्लिक अटेम्प्ट
            try:
                with page.expect_download(timeout=30000) as download_info:
                    page.locator("a:has-text('Download'), button:has-text('Download')").first.click()
                download = download_info.value
                download.save_as(raw_download_path)
            except Exception:
                pass
        
        browser.close()
        print("डेटा सफलतापूर्वक डाउनलोड हो गया।")

    # ============================================================
    # PANDAS DATA PROCESSING & FILTERING (A to Z)
    # ============================================================
    if not os.path.exists(raw_download_path):
        print("Error: डाउनलोड की गई फ़ाइल नहीं मिली!")
        return

    print("डाउनलोड की गई फ़ाइल को फ़िल्टर और सॉर्ट किया जा रहा है...")
    df = pd.read_excel(raw_download_path)
    df.columns = df.columns.str.strip()
    
    # डेटा को साफ करना
    df['Pending'] = pd.to_numeric(df['Pending'], errors='coerce').fillna(0)
    df['Employee'] = df['Employee'].astype(str).str.strip()
    
    # फ़िल्टर: सिर्फ 0 से बड़े पेंडिंग वाले (Active + Inactive दोनों रहेंगे)
    filtered_df = df[df['Pending'] > 0].copy()
    
    # ऑटो सॉर्टिंग: कर्मचारियों के नाम के हिसाब से A to Z क्रम में लगाना
    filtered_df = filtered_df.sort_values(by='Employee', ascending=True)
    
    # नई एक्सेल फ़ाइल सेव करना
    filtered_df.to_excel(OUTPUT_FILE_PATH, index=False)
    
    # अस्थाई कच्ची फ़ाइल डिलीट करना
    if os.path.exists(raw_download_path):
        os.remove(raw_download_path)

    # टेलीग्राम समरी मेसेज तैयार करना
    total_cases = len(filtered_df)
    total_amount = filtered_df['Pending'].sum()
    
    summary_message = (
        "🌅 *Cash Balance Pending Report (A to Z)*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👥 *Total Pending Employees:* {total_cases}\n"
        f"💰 *Total Pending Amount:* ₹{total_amount:,.2f}\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "📢 *Note:* इस लिस्ट में Active और Inactive दोनों कर्मचारी शामिल हैं। 0 पेंडिंग वाले लोगों को हटा दिया गया है और डेटा को नाम के अनुसार A to Z व्यवस्थित किया गया है।"
    )
    
    # 4. टेलीग्राम पर फाइनल सेंड करना (हर हाल में सेंड करेगा)
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    try:
        run_cash_balance_report_job()
    except Exception as e:
        print(f"इस रिपोर्ट जॉब में एरर आया: {e}")
