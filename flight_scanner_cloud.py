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
# ⚙️ قراءة المتغيرات السحابية مع التنظيف الذاتي
# ==========================================
def clean_env(key, default=""):
    val = os.environ.get(key, default)
    if val:
        return val.strip("[]'\" \t\r\n")
    return default

TELEGRAM_BOT_TOKEN = clean_env("TELEGRAM_BOT_TOKEN", "8572404205:AAHYoKETrHLjG_lUMpcTrFbB0hNLjqPbDJ0")
TELEGRAM_CHAT_ID = clean_env("TELEGRAM_CHAT_ID", "536683079")
GOOGLE_SHEET_WEBHOOK_URL = clean_env("GOOGLE_SHEET_WEBHOOK_URL", "https://script.google.com/macros/s/AKfycbwglVr2b3S7C97mMKKL2pJNct_yO3R10Fz0a3JCBsbYxtax56-tz-7_8SFwh6RubIdQJw/exec")
GOOGLE_SHEET_VIEW_URL = clean_env("GOOGLE_SHEET_VIEW_URL", "https://docs.google.com/spreadsheets/d/1eozILOpDIk3KHVIyqJovDIIXTaMAM0cXpeb-Czerr9I/edit?gid=0#gid=0")

HISTORY_FILE = "flight_price_history.json"
EXCEL_FILE = "google_flights_weekends.xlsx"
CHART_IMAGE_PATH = "flight_price_chart.png"

BAGGAGE_FEES = {
    "طيران أديل": 140,
    "طيران ناس": 150,
    "الخطوط السعودية": 0
}


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
        print(f"⚠️ فشل حفظ السجل التاريخي: {e}")


def send_telegram_msg(message, reply_markup=None, high_priority=False):
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        params = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": message,
            "parse_mode": "HTML",
            "disable_web_page_preview": "true",
            "disable_notification": "false" if high_priority else "true"
        }
        if reply_markup:
            params["reply_markup"] = json.dumps(reply_markup)

        payload = urllib.parse.urlencode(params).encode("utf-8")
        req = urllib.request.Request(url, data=payload)
        urllib.request.urlopen(req, timeout=15)
    except Exception as e:
        print(f"⚠️ خطأ إرسال رسالة التيليجرام: {e}")


def send_telegram_photo(photo_path, caption=""):
    try:
        boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
        with open(photo_path, 'rb') as f:
            file_data = f.read()

        parts = [
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"chat_id\"\r\n\r\n{TELEGRAM_CHAT_ID}\r\n".encode('utf-8'),
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"caption\"\r\n\r\n{caption}\r\n".encode('utf-8'),
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"parse_mode\"\r\n\r\nHTML\r\n".encode('utf-8'),
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"photo\"; filename=\"chart.png\"\r\nContent-Type: image/png\r\n\r\n".encode('utf-8') +
            file_data + f"\r\n--{boundary}--\r\n".encode('utf-8')
        ]

        body = b''.join(parts)
        req = urllib.request.Request(
            url,
            data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
        )
        urllib.request.urlopen(req, timeout=25)
    except Exception as e:
        print(f"⚠️ فشل إرسال صورة الرسم البياني: {e}")


def sync_to_google_sheets(results):
    if not GOOGLE_SHEET_WEBHOOK_URL.startswith("http"):
        print(f"⚠️ رابط الويب هوك غير صالح: {GOOGLE_SHEET_WEBHOOK_URL}")
        return

    ksa_time = (datetime.datetime.utcnow() + datetime.timedelta(hours=3)).strftime("%Y-%m-%d %I:%M %p")
    payload_data = {
        "updated_at": f"{ksa_time} (توقيت مكة)",
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
        with urllib.request.urlopen(req, timeout=25) as resp:
            print(f"📤 تم تحديث الشيت السحابي: {resp.read().decode('utf-8')[:150]}")
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
        if weekday == 3:  # الخميس
            fri = current + datetime.timedelta(days=1)
            sat = current + datetime.timedelta(days=2)
            if fri <= end_date:
                pairs.append((current.strftime('%Y-%m-%d'), fri.strftime('%Y-%m-%d'), "خميس إلى جمعة"))
            if sat <= end_date:
                pairs.append((current.strftime('%Y-%m-%d'), sat.strftime('%Y-%m-%d'), "خميس إلى سبت"))
        elif weekday == 4:  # الجمعة
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
    # 1. استخراج وقت الإقلاع
    time_ar = ""
    raw_time = ""
    m_time = re.search(r'(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)\s*[\u2013\-–]\s*(\d{1,2}:\d{2})', text)
    if m_time:
        raw_time = m_time.group(1).strip()
        time_ar = raw_time.replace("AM", "ص").replace("PM", "م").replace("am", "ص").replace("pm", "م")

    # 2. خطوة الأمان الحيوية: شطب جميع التوقيتات والمدد وأوزان الحقائب من النص تماماً
    # هذا يمنع نهائياً قراءة الدقيقة "25" من الساعة "8:25" كسعر!
    cleaned = re.sub(r'\d{1,2}:\d{2}(?:\s*(?:AM|PM|am|pm))?', '', text)
    cleaned = re.sub(r'\d+\s*(?:hr|h|min|m|ساعة|دقيقة|kg|كجم|co2e?).*', '', cleaned, flags=re.IGNORECASE)

    price = None
    # 3. قبول السعر فقط إذا كان مقترناً حصراً برمز العملة (SAR / ر.س / ريال)
    matches = re.findall(r'(?:sar|ر\.س|ريال|sr)\s*([\d,]+)|([\d,]+)\s*(?:sar|ر\.س|ريال|sr)', cleaned, re.IGNORECASE)
    for m1, m2 in matches:
        raw_val = m1 if m1 else m2
        v = int(raw_val.replace(',', ''))
        # حد الأمان المنطقي لتذكرة ذهاب وعودة إجمالية: بين 130 و 4000 ر.س
        if 130 <= v <= 4000:
            price = v
            break

    airline = parse_airline_name(text)
    return price, airline, time_ar, raw_time


def generate_price_chart(items):
    if not items:
        return False

    try:
        chronological = sorted(items, key=lambda x: str(x.get("تاريخ الذهاب", "")))

        dates = []
        prices = []
        colors = []
        color_map = {
            "طيران أديل": "#10b981",
            "الخطوط السعودية": "#1e3a8a",
            "طيران ناس": "#f59e0b"
        }

        for it in chronological:
            d_str = it.get("تاريخ الذهاب", "")
            short_date = "/".join(d_str.split("-")[1:]) if "-" in d_str else d_str
            dates.append(short_date)
            p = it.get("السعر", 0)
            prices.append(p)
            air = it.get("الناقل", "أخرى")
            c = "#8b5cf6" if "دمج ذكي" in air else color_map.get(air, "#6b7280")
            colors.append(c)

        plt.figure(figsize=(11, 5.5), facecolor='#f8fafc')
        ax = plt.subplot(111)
        ax.set_facecolor('#ffffff')

        plt.plot(range(len(dates)), prices, color='#cbd5e1', linestyle='--', linewidth=1.5, zorder=1)

        for i in range(len(dates)):
            plt.scatter(i, prices[i], color=colors[i], s=75, zorder=3, edgecolors='white', linewidth=1.2)

        target = 450
        plt.axhline(y=target, color='#ef4444', linestyle=':', linewidth=1.5, label=f'Target ({target} SAR)')

        min_p = min(prices)
        min_idx = prices.index(min_p)
        plt.scatter(min_idx, min_p, color='#e11d48', s=160, marker='*', zorder=4, label=f'Lowest ({min_p} SAR)')

        plt.title('ELQ <-> JED Flight Price Trend (Direct Flights)', fontsize=14, fontweight='bold', pad=15, color='#0f172a')
        plt.xlabel('Departure Date (MM/DD)', fontsize=10, fontweight='bold', labelpad=10, color='#475569')
        plt.ylabel('Total Roundtrip Price (SAR)', fontsize=10, fontweight='bold', labelpad=10, color='#475569')

        plt.xticks(range(len(dates)), dates, rotation=55, fontsize=8, color='#334155')
        plt.yticks(fontsize=9, color='#334155')
        plt.grid(True, linestyle=':', alpha=0.5, color='#94a3b8')

        custom_legend = [
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='#10b981', markersize=8, label='Flyadeal'),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='#1e3a8a', markersize=8, label='Saudia'),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='#f59e0b', markersize=8, label='Flynas'),
            plt.Line2D([0], [0], color='#ef4444', linestyle=':', label=f'Target ({target} SAR)'),
            plt.Line2D([0], [0], marker='*', color='w', markerfacecolor='#e11d48', markersize=11, label=f'Lowest ({min_p} SAR)')
        ]
        plt.legend(handles=custom_legend, loc='upper right', framealpha=0.9, fontsize=8)

        plt.tight_layout()
        plt.savefig(CHART_IMAGE_PATH, dpi=160)
        plt.close()
        return True
    except Exception as e:
        print(f"⚠️ خطأ بالرسم البياني: {e}")
        return False


def build_briefing_message(items):
    if not items:
        return "لا توجد رحلات مباشرة مرصودة حالياً."

    today = datetime.date.today()
    upcoming = []
    for x in items:
        try:
            d = datetime.datetime.strptime(x["تاريخ الذهاب"], "%Y-%m-%d").date()
            if d >= today and "🇸🇦" not in x.get("نوع العطلة", ""):
                upcoming.append((d, x))
        except Exception:
            pass
    upcoming = sorted(upcoming, key=lambda t: t[0])
    next_weekend = upcoming[0][1] if upcoming else items[0]
    cheapest_overall = items[0]

    items_with_bags = sorted(
        items,
        key=lambda it: it["السعر"] + BAGGAGE_FEES.get(it["الناقل"], 140)
    )
    best_with_bag = items_with_bags[0]
    best_bag_price = best_with_bag["السعر"] + BAGGAGE_FEES.get(best_with_bag["الناقل"], 140)

    min_p = cheapest_overall["السعر"]
    if min_p <= 360:
        advice = "🟢 <b>توصية الرادار:</b> الأسعار حالياً في <b>القاع التاريخي (358 ر.س أو أقل)</b>. فرصة ممتازة لتأكيد الحجوزات الآن."
    elif min_p <= 420:
        advice = "🟡 <b>توصية الرادار:</b> الأسعار معتدلة ومناسبة للحجز."
    else:
        advice = "🔴 <b>توصية الرادار:</b> الأسعار مرتفعة، ينصح بالانتظار."

    t_next = f" | ⏰ {next_weekend['وقت الإقلاع']}" if next_weekend.get("وقت الإقلاع") else ""
    t_cheap = f" | ⏰ {cheapest_overall['وقت الإقلاع']}" if cheapest_overall.get("وقت الإقلاع") else ""

    ksa_now = datetime.datetime.utcnow() + datetime.timedelta(hours=3)
    text = (
        f"☀️ <b>النشرة الذكية لأسعار طيران (القصيم ⇄ جدة)</b>\n"
        f"📅 <i>{ksa_now.strftime('%A, %d %B %Y')}</i>\n\n"
        f"📍 <b>1. أقرب عطلة نهاية أسبوع:</b>\n"
        f"   🗓 <b>{next_weekend['نوع العطلة']}</b>{t_next}\n"
        f"   🛫 ذهاب: <code>{next_weekend['تاريخ الذهاب']}</code> ⬅ عودة: <code>{next_weekend['تاريخ العودة']}</code>\n"
        f"   💰 <b>السعر الإجمالي:</b> <b>{next_weekend['السعر']} ر.س</b> — ✈️ {next_weekend['الناقل']}\n"
        f"   🔗 <a href='{next_weekend['الرابط']}'>حجز أقرب رحلة ↗</a>\n\n"
        f"🎒 <b>2. أرخص تذكرة خفيفة (بدون شحن):</b>\n"
        f"   💰 <b>{cheapest_overall['السعر']} ر.س إجمالي</b> ({cheapest_overall['نوع العطلة']})\n"
        f"   ✈️ {cheapest_overall['الناقل']}{t_cheap}\n"
        f"   🔗 <a href='{cheapest_overall['الرابط']}'>حجز تذكرة القاع ↗</a>\n\n"
        f"🧳 <b>3. أفضل صفقة شاملة حقيبة شحن (20kg):</b>\n"
        f"   💰 <b>{best_bag_price} ر.س إجمالي</b> — ✈️ {best_with_bag['الناقل']}\n"
        f"   🗓 {best_with_bag['نوع العطلة']} (<code>{best_with_bag['تاريخ الذهاب']}</code>)\n\n"
        f"{advice}\n\n"
        f"📊 <a href='{GOOGLE_SHEET_VIEW_URL}'>فتح جدول Google Sheets المباشر والأرشيف</a>"
    )
    return text


def run_cloud_scan():
    print("☁️ بدء تشغيل الرادار السحابي عبر GitHub Actions 24/7...")
    pairs = get_all_monitored_pairs(months_ahead=3)
    results = []
    history = load_history()

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            locale="en-US",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 768}
        )

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
                print(f"[{idx}/{len(pairs)}] ✅ رصد حقيقي: {trip_type} -> {cur_p} ر.س ({air})")

                # صائد العروض الخاطفة الحقيقي (خصم حقيقي لا يقل عن 70 ر.س أو قاع حقيقي تحت 260 ر.س)
                prev_p = history.get(flight_key)
                is_deep_drop = prev_p and (prev_p - cur_p >= 70)
                is_rock_bottom = cur_p <= 260

                if is_deep_drop or is_rock_bottom:
                    reason = f"📉 هبوط حاد بمقدار {prev_p - cur_p} ر.س!" if is_deep_drop else "🔥 كسر سعر القاع التاريخي (أقل من 260 ر.س)!"
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

            time.sleep(random.uniform(0.5, 1.2))

        browser.close()

    save_history(history)

    if results:
        results = sorted(results, key=lambda x: x["السعر"])
        sync_to_google_sheets(results)
        df = pd.DataFrame(results)

        ksa_now = datetime.datetime.utcnow() + datetime.timedelta(hours=3)
        is_manual_trigger = os.environ.get("GITHUB_EVENT_NAME") == "workflow_dispatch"
        is_morning_window = (8 <= ksa_now.hour < 10)

        if is_manual_trigger or is_morning_window:
            print("📢 إرسال التقرير الشامل والمخطط البياني للتيليجرام...")
            generate_price_chart(results)
            briefing_text = build_briefing_message(results)

            if os.path.exists(CHART_IMAGE_PATH):
                caption = f"📈 <b>مخطط حركة أسعار (القصيم ⇄ جدة)</b>\n• أديل (أخضر) | السعودية (كحلي) | ناس (ذهبي)"
                send_telegram_photo(CHART_IMAGE_PATH, caption=caption)

            send_telegram_msg(briefing_text)
    else:
        df = pd.DataFrame(columns=["نوع العطلة", "تاريخ الذهاب", "وقت الإقلاع", "تاريخ العودة", "الناقل", "السعر", "الرابط"])

    df.to_excel(EXCEL_FILE, index=False)
    print(f"✅ اكتملت الدورة السحابية بنجاح! تم رصد {len(results)} رحلة حقيقية.")


if __name__ == "__main__":
    run_cloud_scan()
