import calendar
import datetime
import json
import os
import random
import re
import time
import urllib.parse
import urllib.request
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from playwright.sync_api import sync_playwright

# ==========================================
# ⚙️ قراءة المتغيرات من خزنة GitHub Secrets
# ==========================================
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
GOOGLE_SHEET_WEBHOOK_URL = os.environ.get("GOOGLE_SHEET_WEBHOOK_URL", "")
GOOGLE_SHEET_VIEW_URL = os.environ.get("GOOGLE_SHEET_VIEW_URL", "")

HISTORY_FILE = "flight_price_history.json"
EXCEL_FILE = "google_flights_weekends.xlsx"
CHART_IMAGE_PATH = "flight_price_chart.png"

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_history(hist):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(hist, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ فشل حفظ السجل: {e}")

def send_telegram_msg(message, high_priority=False):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        params = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": message,
            "parse_mode": "HTML",
            "disable_web_page_preview": "true",
            "disable_notification": "false" if high_priority else "true"
        }
        payload = urllib.parse.urlencode(params).encode("utf-8")
        req = urllib.request.Request(url, data=payload)
        urllib.request.urlopen(req, timeout=15)
    except Exception as e:
        print(f"⚠️ خطأ إرسال التيليجرام: {e}")

def sync_to_google_sheets(results):
    if not GOOGLE_SHEET_WEBHOOK_URL.startswith("http"):
        print("⚠️ رابط Webhook غير مضبوط")
        return

    payload_data = {
        "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p"),
        "flights": [
            {
                "trip_type": item["نوع العطلة"],
                "dep": f"{item['تاريخ الذهاب']} ({item['وقت الإقلاع']})" if item.get("وقت الإقلاع") else item["تاريخ الذهاب"],
                "ret": item["تاريخ العودة"],
                "airline": item["الناقل"],
                "price": item["السعر"],
                "link": item["الرابط"]
            }
            for item in results
        ]
    }

    try:
        payload_bytes = json.dumps(payload_data).encode("utf-8")
        req = urllib.request.Request(
            GOOGLE_SHEET_WEBHOOK_URL,
            data=payload_bytes,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            print("📤 تم إرسال وتحديث البيانات في Google Sheets بنجاح!")
    except Exception as e:
        print(f"⚠️ خطأ رفع الشيت: {e}")

def get_all_monitored_pairs(months_ahead=3):
    today = datetime.date.today()
    end_date = today + datetime.timedelta(days=months_ahead * 30)
    pairs = []

    special_events = [
        ("2026-09-22", "2026-09-26", "🇸🇦 إجازة اليوم الوطني (ثلاثاء إلى سبت)"),
        ("2026-09-23", "2026-09-26", "🇸🇦 إجازة اليوم الوطني (أربعاء إلى سبت)"),
        ("2026-09-23", "2026-09-25", "🇸🇦 إجازة اليوم الوطني (أربعاء إلى جمعة)"),
        ("2026-10-21", "2026-10-24", "🏖️ عطلة مطولة (أربعاء إلى سبت)"),
    ]

    for dep_str, ret_str, label in special_events:
        d_dep = datetime.datetime.strptime(dep_str, "%Y-%m-%d").date()
        d_ret = datetime.datetime.strptime(ret_str, "%Y-%m-%d").date()
        if today < d_dep and d_ret <= end_date:
            pairs.append((dep_str, ret_str, label))

    current = today + datetime.timedelta(days=1)
    while current <= end_date:
        weekday = current.weekday()
        if weekday == 3:
            fri = current + datetime.timedelta(days=1)
            sat = current + datetime.timedelta(days=2)
            if fri <= end_date:
                pairs.append((current.strftime('%Y-%m-%d'), fri.strftime('%Y-%m-%d'), "خميس إلى جمعة"))
            if sat <= end_date:
                pairs.append((current.strftime('%Y-%m-%d'), sat.strftime('%Y-%m-%d'), "خميس إلى سبت"))
        elif weekday == 4:
            sat = current + datetime.timedelta(days=1)
            if sat <= end_date:
                pairs.append((current.strftime('%Y-%m-%d'), sat.strftime('%Y-%m-%d'), "جمعة إلى سبت"))

        current += datetime.timedelta(days=1)

    return pairs

def parse_airline_name(text):
    t = text.lower()
    found = []
    if "flyadeal" in t or "أديل" in t:
        found.append("طيران أديل")
    if "flynas" in t or "ناس" in t:
        found.append("طيران ناس")
    if "saudia" in t or "السعودية" in t:
        found.append("الخطوط السعودية")

    if len(found) > 1:
        return " + ".join(found) + " (دمج ذكي 🔀)"
    elif len(found) == 1:
        return found[0]
    return "رحلة مباشرة"

def parse_flight_card_text(text):
    time_ar = ""
    raw_time = ""
    m_time = re.search(r'(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)\s*[\u2013\-–]\s*(\d{1,2}:\d{2})', text)
    if m_time:
        raw_time = m_time.group(1).strip()
        time_ar = raw_time.replace("AM", "ص").replace("PM", "م").replace("am", "ص").replace("pm", "م")

    cleaned = re.sub(r'\d+\s*(?:kg|كجم|co2e?).*', '', text, flags=re.IGNORECASE)

    price = None
    matches = re.findall(r'(?:sar|ر\.س|ريال)\s*([\d,]+)|([\d,]+)\s*(?:sar|ر\.س|ريال)', cleaned, re.IGNORECASE)
    for m1, m2 in matches:
        raw_val = m1 if m1 else m2
        v = int(raw_val.replace(',', ''))
        if 150 <= v <= 4000:
            price = v
            break

    if not price:
        alt_matches = re.findall(r'\b(\d{3,4})\b', cleaned)
        for val in alt_matches:
            v = int(val)
            if 200 <= v <= 3500:
                price = v
                break

    airline = parse_airline_name(text)
    return price, airline, time_ar, raw_time

def run_cloud_scan():
    print("☁️ بدء فحص الرادار السحابي عبر GitHub Actions...")
    pairs = get_all_monitored_pairs(months_ahead=3)
    results = []
    history = load_history()

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        context = browser.new_context(
            locale="en-US",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 768}
        )
        
        # تخطي شاشة كوكيز جوجل التلقائية
        context.add_cookies([
            {"name": "SOCS", "value": "CAESHAgBEhJnd3NfMjAyNDA4MDctMF9SQzIaAmVuIAEaBgiA_L20Bg", "domain": ".google.com", "path": "/"},
            {"name": "CONSENT", "value": "PENDING+999", "domain": ".google.com", "path": "/"}
        ])
        
        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        page = context.new_page()

        for idx, (dep, ret, trip_type) in enumerate(pairs, 1):
            url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep}%20through%20{ret}%20nonstop&curr=SAR&hl=en&gl=sa"
            cheapest_flight = None

            try:
                page.goto(url, wait_until="domcontentloaded", timeout=18000)

                # تخطي أي شاشة موافقة طارئة
                if "consent.google.com" in page.url:
                    try:
                        page.locator('button:has-text("Accept all"), button:has-text("I agree")').first.click(timeout=3000)
                        page.wait_for_load_state("domcontentloaded", timeout=6000)
                    except Exception:
                        pass

                try:
                    page.wait_for_selector("li.pIav2d, div.pIav2d", timeout=4500)
                except Exception:
                    time.sleep(1.5)

                cards = page.locator("li.pIav2d, div.pIav2d").all()
                nonstop_options = []

                for c in cards:
                    txt = c.inner_text()
                    # دعم كل مسميات الطيران المباشر (إنجليزي وعربي)
                    if "nonstop" in txt.lower() or "مباشر" in txt or "بدون توقف" in txt:
                        price, airline, time_ar, raw_time = parse_flight_card_text(txt)
                        if price:
                            nonstop_options.append({
                                "price": price,
                                "airline": airline,
                                "time_ar": time_ar,
                                "raw_time": raw_time
                            })

                if nonstop_options:
                    if "جمعة إلى سبت" in trip_type:
                        pm_options = [o for o in nonstop_options if "PM" in o["raw_time"].upper() or "م" in o["time_ar"]]
                        cheapest_flight = min(pm_options, key=lambda x: x["price"]) if pm_options else min(nonstop_options, key=lambda x: x["price"])
                    else:
                        cheapest_flight = min(nonstop_options, key=lambda x: x["price"])

            except Exception:
                pass

            if cheapest_flight:
                cur_p = cheapest_flight["price"]
                air = cheapest_flight["airline"]
                f_time = cheapest_flight["time_ar"]
                flight_key = f"{dep}_{ret}"
                print(f"[{idx}/{len(pairs)}] ✅ رصد: {trip_type} -> {cur_p} ر.س ({air})")

                prev_p = history.get(flight_key)
                is_deep_drop = prev_p and (prev_p - cur_p >= 70)
                is_rock_bottom = cur_p <= 310

                if is_deep_drop or is_rock_bottom:
                    reason = f"📉 هبوط بمقدار {prev_p - cur_p} ر.س!" if is_deep_drop else "🔥 كسر سعر القاع التاريخي!"
                    flash_msg = (
                        f"⚡🚨 <b>عرض ترويجي خاطف (رصد سحابي 24/7)</b>\n\n"
                        f"🗓 <b>{trip_type}</b>\n"
                        f"🛫 <b>الذهاب:</b> <code>{dep}</code> (⏰ {f_time})\n"
                        f"🛬 <b>العودة:</b> <code>{ret}</code>\n"
                        f"✈️ <b>الناقل:</b> {air}\n"
                        f"💰 <b>السعر الإجمالي:</b> <b>{cur_p} ر.س فقط!</b>\n"
                        f"📢 <b>السبب:</b> {reason}\n\n"
                        f"🔗 <a href='{url}'>اضغط هنا لحجز العرض فوراً ↗</a>"
                    )
                    send_telegram_msg(flash_msg, high_priority=True)

                history[flight_key] = cur_p
                results.append({
                    "نوع العطلة": trip_type,
                    "تاريخ الذهاب": dep,
                    "وقت الإقلاع": f_time,
                    "تاريخ العودة": ret,
                    "الناقل": air,
                    "السعر": cur_p,
                    "الرابط": url
                })
            else:
                print(f"[{idx}/{len(pairs)}] ℹ️ تم فحص الرابط: {trip_type}")

            time.sleep(random.uniform(0.4, 0.9))

        browser.close()

    save_history(history)

    if results:
        results = sorted(results, key=lambda x: x["السعر"])
        sync_to_google_sheets(results)
        df = pd.DataFrame(results)
    else:
        df = pd.DataFrame(columns=["نوع العطلة", "تاريخ الذهاب", "وقت الإقلاع", "تاريخ العودة", "الناقل", "السعر", "الرابط"])

    df.to_excel(EXCEL_FILE, index=False)
    print(f"✅ اكتمل الفحص بنجاح! تم رصد {len(results)} رحلة وتحديث Google Sheets.")

if __name__ == "__main__":
    run_cloud_scan()
