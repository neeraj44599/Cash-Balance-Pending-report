from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from datetime import datetime
import os
import time
import re
import pandas as pd
import requests

# ============================================================
# CONFIGURATION
# ============================================================
USERNAME = os.getenv("ACEN_USERNAME")
PASSWORD = os.getenv("ACEN_PASSWORD")
LOGIN_URL = os.getenv("ACEN_LOGIN_URL")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

BASE_DIR = os.getcwd()
OUTPUT_FILE_PATH = os.path.join(BASE_DIR, "Cash_Balance_Pending_Report.xlsx")

# ============================================================
# TELEGRAM FUNCTION (FIXED API LINK)
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
        
        print("लॉगिन पेज ओपन किया जा रहा है...")
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
        
        # 💡 iframe को हैंडल करना: पोर्टल के अंदर सही फ़्रेम को टारगेट करें
        # (अगर मुख्य पेज पर लॉगिन है तो यह सीधे चलेगा, नहीं तो फ़्रेम के अंदर इनपुट्स ढूँढेगा)
        target = page
        time.sleep(3)
        
        frames = page.frames
        if len(frames) > 1:
            print(f"वेबसाइट पर {len(frames)} iframes मिले। सही फ़्रेम सेलेक्ट किया जा रहा है...")
            for f in frames:
                if f.name or f.url:
                    target = f
                    break

        # 1. क्रेडेंशियल्स भरना और लॉगिन करना
        try:
            print("क्रेडेंशियल्स भरे जा रहे हैं...")
            target.locator("input[type='text'], input[type='email'], input[name='username'], #username").first.fill(USERNAME)
            target.locator("input[type='password'], #password").first.fill(PASSWORD)
            
            submit_btn = target.locator("button[type='submit'], input[type='submit'], button:has-text('Login'), #login-btn").first
            submit_btn.click()
            print("लॉगिन बटन पर क्लिक कर दिया गया है।")
            time.sleep(5)
        except Exception as e:
            print(f"लॉगिन फ़ील्ड भरने में रुकावट (शायद ऑटो-लॉगिन एक्टिव था): {e}")

        # दोबारा फ़्रेम चेक करें (लॉगिन के बाद मुख्य डेटा अक्सर नए iframe में आता है)
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        target = page
        for f in page.frames:
            if "main" in f.name.lower() or "report" in f.name.lower() or len(f.name) > 0:
                target = f
                break

        # 2. Reports बटन पर जाना (नया और मजबूत फ़्रेम-ओरिएंटेड तरीका)
        print("Reports सेक्शन में जाया जा रहा... बटन ढूँढा जा रहा है...")
        try:
            # यह पहले सही फ़्रेम में ढूँढेगा, न मिलने पर मुख्य पेज पर
            reports_locator = target.get_by_text("Reports", exact=True)
            reports_locator.wait_for(state="visible", timeout=60000)
            reports_locator.click()
            print("Reports बटन पर सफलतापूर्वक क्लिक हो गया!")
            time.sleep(3)
        except Exception as e:
            print("मुख्य फ़्रेम में बटन नहीं मिला, बैकअप फ़्रेम सर्च किया जा रहा है...")
            page.locator("a:has-text('Reports'), button:has-text('Reports')").first.click()

        # 3. फ़ाइल डाउनलोड प्रोसेस (Playwright Download handling)
        print("डेटा डाउनलोड ट्रिगर किया जा रहा है...")
        try:
            # डाउनलोड बटन के लिए फ़्रेम को ही टारगेट रखना ज़रूरी है
            download_btn = target.locator("button, a").filter(has_text=re.compile(r"download|export|excel|एक्सेल", re.IGNORECASE)).first
            
            with page.expect_download(timeout=60000) as download_info:
                download_btn.click()
            
            download = download_info.value
            download.save_as(raw_download_path)
            print("डेटा फ़ाइल सफलतापूर्वक डाउनलोड हो गई।")
        except Exception as dl_err:
            print(f"डाउनलोड ट्रिगर फेल हुआ, डिफ़ॉल्ट एक्शन आज़माया जा रहा है: {dl_err}")
            # फॉलबैक: डायरेक्ट क्लिक अटेम्प्ट
            try:
                with page.expect_download(timeout=30000) as download_info:
                    page.locator("a:has-text('Download'), button:has-text('Download')").first.click()
                download = download_info.value
                download.save_as(raw_download_path)
            except Exception:
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
    
    # कॉलम नाम ऑटो-डिटेक्ट करना (ताकि इनएक्टिव वाले मिस न हों)
    pending_col = next((c for c in df.columns if 'pending' in c.lower() or 'cash' in c.lower()), 'Pending')
    emp_col = next((c for c in df.columns if 'emp' in c.lower() or 'name' in c.lower()), 'Employee')
    
    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[emp_col] = df[emp_col].astype(str).str.strip()
    
    # फ़िल्टर: सिर्फ 0 से बड़े पेंडिंग वाले (Active + Inactive दोनों शामिल रहेंगे)
    filtered_df = df[df[pending_col] > 0].copy()
    
    # ऑटो सॉर्टिंग: कर्मचारियों के नाम के हिसाब से A to Z क्रम में लगाना
    filtered_df = filtered_df.sort_values(by=emp_col, ascending=True)
    
    # नई एक्सेल फ़ाइल सेव करना
    filtered_df.to_excel(OUTPUT_FILE_PATH, index=False)
    
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
        "📢 *Note:* इस लिस्ट में Active और Inactive दोनों कर्मचारी शामिल हैं। 0 पेंडिंग वाले लोगों को हटा दिया गया है और डेटा को नाम के अनुसार A to Z व्यवस्थित किया गया है।"
    )
    
    # फाइनल सेंड
    send_telegram_report(summary_message, OUTPUT_FILE_PATH)

if __name__ == "__main__":
    run_cash_balance_report_job()
