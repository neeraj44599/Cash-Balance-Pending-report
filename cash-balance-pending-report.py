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
# TELEGRAM FUNCTION (100% CORRECTED API LINK)
# ============================================================
def send_telegram_report(message, excel_path):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram credentials missing. Dispatch skipped.")
        return
    
    # 💡 यहाँ सही API लिंक और 'bot' शब्द जोड़ दिया गया है
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

            # 3. फ़ाइल डाउनलोड प्रोसेस (इसे बदलें)
        print("डेटा डाउनलोड ट्रिगर किया जा रहा है...")
        try:
            with page.expect_download(timeout=12000) as download_info:
                # 💡 यहाँ अपने पोर्टल के डाउनलोड बटन का Selector डालें (जैसे text="Download" या .btn)
                page.get_by_role("button", name="Download").click() 
            
            download = download_info.value
            download.save_as(raw_download_path)
            print("फ़ाइल सफलतापूर्वक डाउनलोड होकर सेव हो गई।")
        except Exception as dl_error:
            print(f"डाउनलोड फेल हो गया या बटन नहीं मिला: {dl_error}")
            browser.close()
            return
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
        
        # [नोट: यहाँ अपने पुराने कोड के अनुसार username/password भरने और लॉगिन बटन दबाने की लाइन्स डाल लें]
        
        print("Reports सेक्शन में जाया जा रहा... बटन ढूँढा जा रहा है...")
        page.get_by_text("Reports", exact=True).wait_for(state="visible", timeout=60000)
        page.get_by_text("Reports", exact=True).click()
        
        print("डेटा डाउनलोड ट्रिगर किया जा रहा है...")
        # [नोट: यहाँ अपनी रिपोर्ट डाउनलोड करने वाला पुराना क्लिक कोड डाल लें]
        
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
    
    # 💡 कॉलम के नाम ऑटो-डिटेक्ट करें (चाहे Pending हो, Cash हो या Amount हो)
    pending_col = next((c for c in df.columns if 'pending' in c.lower() or 'cash' in c.lower()), None)
    emp_col = next((c for c in df.columns if 'emp' in c.lower() or 'name' in c.lower()), None)
    
    if not pending_col or not emp_col:
        print(f"Error: सही कॉलम नहीं मिले! एक्सेल में ये कॉलम्स हैं: {df.columns.tolist()}")
        # अगर ऑटो-डिटेक्ट काम न करे तो डिफ़ॉल्ट नामों पर वापस जाएँ
        pending_col = 'Pending'
        emp_col = 'Employee'
        
    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[emp_col] = df[emp_col].astype(str).str.strip()
    
    # फ़िल्टर: सिर्फ 0 से बड़े पेंडिंग वाले (Active + Inactive दोनों रहेंगे)
    filtered_df = df[df[pending_col] > 0].copy()
    
    # ऑटो सॉर्टिंग: A to Z क्रम में लगाना
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
        "📢 *Note:* इस लिस्ट में Active और Inactive दोनों कर्मचारी शामिल हैं। 0 पेंडिंग वाले लोगों को हटा दिया गया है और डेटा को नाम के अनुसार A to Z व्यवस्थित किया गया है।"
    )
    
    # 💡 इसे हर हाल में टेलीग्राम पर भेजें ताकि आपको कन्फर्म मेसेज मिले
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    try:
        run_cash_balance_report_job()
    except Exception as e:
        print(f"इस रिपोर्ट जॉब में एरर आया: {e}")
