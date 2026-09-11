import calendar
import datetime
import json
import os
import random
import re
import time
import traceback
import urllib.parse
import urllib.request
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from playwright.sync_api import sync_playwright

# ==========================================
# ⚙️ قراءة المتغيرات السحابية وتنظيفها
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

CUSTOM_DEP = clean_env("CUSTOM_DEP", "")
CUSTOM_RET = clean_env("CUSTOM_RET", "")

HISTORY_FILE = "flight_price_history.json"
EXCEL_FILE = "google_flights_weekends.xlsx"
CHART_IMAGE_PATH = "flight_price_chart.png"

BAGGAGE_FEES = {
    "طيران أديل": 140,
    "طيران ناس": 150,
    "الخطوط السعودية": 0
}

# ==========================================
# 🛡️ مصفوفة بصمات التصفح للتخفي ومقاومة الحظر
# ==========================================
REALISTIC_PROFILES = [
    {
        "ua": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
        "viewport": {"width": 1440, "height": 900},
        "platform": "Win32",
        "vendor": "Google Inc.",
        "renderer": "ANGLE (Intel, Intel(R) Iris(R) Xe Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)"
    },
    {
        "ua": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 Edg/128.0.0.0",
        "viewport": {"width": 1366, "height": 768},
        "platform": "Win32",
        "vendor": "Google Inc.",
        "renderer": "ANGLE (NVIDIA, NVIDIA GeForce RTX 3060 Direct3D11 vs_5_0 ps_5_0, D3D11)"
    },
    {
        "ua": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
        "viewport": {"width": 1536, "height": 864},
        "platform": "MacIntel",
        "vendor": "Apple Computer, Inc.",
        "renderer": "Apple M2 Pro"
    },
    {
        "ua": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0",
        "viewport": {"width": 1920, "height": 1080},
        "platform": "Win32",
        "vendor": "",
        "renderer": "AMD Radeon RX 6700 XT"
    }
]


def load_history():
    default_state = {"prices": {}, "alerts": {}}
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    if "prices" not in data:
                        return {"prices": data, "alerts": {}}
                    return data
        except Exception:
            pass
    return default_state


def save_history(hist_data):
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(hist_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ فشل حفظ السجل التاريخي: {e}")


def get_direct_booking_link(airline, dep, ret):
    air = airline or ""
    if "أديل" in air or "flyadeal" in air.lower():
        return f"https://www.flyadeal.com/ar/search-flight?origin=ELQ&destination=JED&departureDate={dep}&returnDate={ret}&adults=1&tripType=RoundTrip"
    elif "ناس" in air or "flynas" in air.lower():
        return f"https://www.flynas.com/ar/booking/flight-search?origin=ELQ&destination=JED&departureDate={dep}&returnDate={ret}&adults=1&tripType=RoundTrip"
    elif "السعودية" in air or "saudia" in air.lower():
        return f"https://www.saudia.com/ar/booking?origin=ELQ&destination=JED&departureDate={dep}&returnDate={ret}&adults=1&tripType=RoundTrip"
    return f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep}%20through%20{ret}%20nonstop&curr=SAR&hl=ar&gl=sa"


def get_google_calendar_link(trip_type, airline, dep, ret, price, url):
    dep_clean = dep.replace("-", "")
    ret_clean = ret.replace("-", "")
    title = urllib.parse.quote(f"✈️ رحلة القصيم ⇄ جدة ({airline})")
    details = urllib.parse.quote(f"🗓 {trip_type}\n💰 السعر الإجمالي: {price} ر.س\n✈️ الناقل: {airline}\n🔗 رابط الحجز: {url}")
    location = urllib.parse.quote("مطار الأمير نايف بن عبدالعزيز (ELQ) ⇄ مطار الملك عبدالعزيز (JED)")
    return f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={title}&dates={dep_clean}/{ret_clean}&details={details}&location={location}"


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


def sync_to_google_sheets(results, execution_time_sec):
    if not GOOGLE_SHEET_WEBHOOK_URL.startswith("http"):
        print(f"⚠️ رابط الويب هوك غير صالح: {GOOGLE_SHEET_WEBHOOK_URL}")
        return

    ksa_now = datetime.datetime.utcnow() + datetime.timedelta(hours=3)
    payload_data = {
        "updated_at": ksa_now.strftime("%Y-%m-%d %I:%M %p"),
        "raw_timestamp": int(ksa_now.timestamp()),
        "scan_duration_sec": execution_time_sec,
        "flight_count": len(results),
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
    time_ar = ""
    raw_time = ""
    m_time = re.search(r'(\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)?)\s*[\u2013\-–]\s*(\d{1,2}:\d{2})', text)
    if m_time:
        raw_time = m_time.group(1).strip()
        time_ar = raw_time.replace("AM", "ص").replace("PM", "م").replace("am", "ص").replace("pm", "م")

    cleaned = re.sub(r'\d{1,2}:\d{2}(?:\s*(?:AM|PM|am|pm))?', '', text)
    cleaned = re.sub(r'\d+\s*(?:hr|h|min|m|ساعة|دقيقة|kg|كجم|co2e?).*', '', cleaned, flags=re.IGNORECASE)

    price = None
    matches = re.findall(r'(?:sar|ر\.س|ريال|sr)\s*([\d,]+)|([\d,]+)\s*(?:sar|ر\.س|ريال|sr)', cleaned, re.IGNORECASE)
    for m1, m2 in matches:
        raw_val = m1 if m1 else m2
        v = int(raw_val.replace(',', ''))
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

    direct_next = get_direct_booking_link(next_weekend['الناقل'], next_weekend['تاريخ الذهاب'], next_weekend['تاريخ العودة'])
    direct_cheap = get_direct_booking_link(cheapest_overall['الناقل'], cheapest_overall['تاريخ الذهاب'], cheapest_overall['تاريخ العودة'])

    ksa_now = datetime.datetime.utcnow() + datetime.timedelta(hours=3)
    text = (
        f"☀️ <b>النشرة الذكية لأسعار طيران (القصيم ⇄ جدة)</b>\n"
        f"📅 <i>{ksa_now.strftime('%A, %d %B %Y')}</i>\n\n"
        f"📍 <b>1. أقرب عطلة نهاية أسبوع:</b>\n"
        f"   🗓 <b>{next_weekend['نوع العطلة']}</b>{t_next}\n"
        f"   🛫 ذهاب: <code>{next_weekend['تاريخ الذهاب']}</code> ⬅ عودة: <code>{next_weekend['تاريخ العودة']}</code>\n"
        f"   💰 <b>السعر الإجمالي:</b> <b>{next_weekend['السعر']} ر.س</b> — ✈️ {next_weekend['الناقل']}\n"
        f"   ✈️ <a href='{direct_next}'>حجز مباشر من موقع {next_weekend['الناقل']} ↗</a>\n\n"
        f"🎒 <b>2. أرخص تذكرة خفيفة (بدون شحن):</b>\n"
        f"   💰 <b>{cheapest_overall['السعر']} ر.س إجمالي</b> ({cheapest_overall['نوع العطلة']})\n"
        f"   ✈️ {cheapest_overall['الناقل']}{t_cheap}\n"
        f"   ✈️ <a href='{direct_cheap}'>حجز تذكرة القاع مباشرة ↗</a>\n\n"
        f"🧳 <b>3. أفضل صفقة شاملة حقيبة شحن (20kg):</b>\n"
        f"   💰 <b>{best_bag_price} ر.س إجمالي</b> — ✈️ {best_with_bag['الناقل']}\n"
        f"   🗓 {best_with_bag['نوع العطلة']} (<code>{best_with_bag['تاريخ الذهاب']}</code>)\n\n"
        f"{advice}\n\n"
        f"📊 <a href='{GOOGLE_SHEET_VIEW_URL}'>فتح جدول Google Sheets المباشر والأرشيف</a>"
    )
    return text


def create_stealth_context(browser):
    """إنشاء سياق تصفح متخفٍ ومحصن بمحاكاة السلوك البشري"""
    profile = random.choice(REALISTIC_PROFILES)
    context = browser.new_context(
        locale="ar-SA",
        timezone_id="Asia/Riyadh",
        user_agent=profile["ua"],
        viewport=profile["viewport"],
        device_scale_factor=1,
        is_mobile=False,
        has_touch=False
    )

    context.add_cookies([
        {"name": "SOCS", "value": "CAESHAgBEhJnd3NfMjAyNDA4MDctMF9SQzIaAmVuIAEaBgiA_L20Bg", "domain": ".google.com", "path": "/"},
        {"name": "CONSENT", "value": "PENDING+999", "domain": ".google.com", "path": "/"}
    ])

    # حقن سكريبتات إخفاء الآلية ومحاكاة العتاد
    stealth_js = f"""
    // 1. إخفاء مؤشر التشغيل الآلي
    Object.defineProperty(navigator, 'webdriver', {{ get: () => undefined }});
    
    // 2. تزييف لغات المتصفح الطبيعية
    Object.defineProperty(navigator, 'languages', {{ get: () => ['ar-SA', 'ar', 'en-US', 'en'] }});
    
    // 3. تزييف كرت الشاشة الحقيقي لمنع كشف SwiftShader
    const getParameter = WebGLRenderingContext.prototype.getParameter;
    WebGLRenderingContext.prototype.getParameter = function(parameter) {{
        if (parameter === 37445) return '{profile["vendor"]}';
        if (parameter === 37446) return '{profile["renderer"]}';
        return getParameter.apply(this, arguments);
    }};

    // 4. محاكاة وجود إضافات المتصفح
    Object.defineProperty(navigator, 'plugins', {{ get: () => [1, 2, 3, 4, 5] }});
    """
    context.add_init_script(stealth_js)
    return context


def run_custom_date_probe(dep, ret):
    print(f"🎯 بدء الفاحص الخاطف للتواريخ الحرة (درع التخفي مفعّل): {dep} ⬅ {ret}")
    url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep}%20through%20{ret}%20nonstop&curr=SAR&hl=en&gl=sa"
    available_flights = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-infobars"
            ]
        )
        context = create_stealth_context(browser)
        page = context.new_page()

        try:
            page.goto(url, wait_until="domcontentloaded", timeout=20000)
            if "consent.google.com" in page.url:
                try:
                    page.locator('button:has-text("Accept all"), button:has-text("I agree")').first.click(timeout=3000)
                    page.wait_for_load_state("domcontentloaded", timeout=5000)
                except Exception:
                    pass

            # تمرير خفيف لمحاكاة التفاعل الطبيعي
            page.mouse.wheel(0, random.randint(150, 300))
            try:
                page.wait_for_selector("li.pIav2d, div.pIav2d", timeout=6000)
            except Exception:
                time.sleep(2)

            cards = page.locator("li.pIav2d, div.pIav2d").all()
            for c in cards:
                txt = c.inner_text()
                if "nonstop" in txt.lower() or "مباشر" in txt or "بدون توقف" in txt:
                    price, airline, time_ar, raw_time = parse_flight_card_text(txt)
                    if price:
                        available_flights.append({
                            "price": price,
                            "airline": airline,
                            "time_ar": time_ar,
                            "raw_time": raw_time
                        })
        except Exception as e:
            print(f"⚠️ خطأ أثناء الفحص المخصص: {e}")
        finally:
            browser.close()

    if not available_flights:
        msg = (
            f"🔍 <b>تقرير الفحص المخصص (القصيم ⇄ جدة)</b>\n\n"
            f"🛫 <b>الذهاب:</b> <code>{dep}</code>\n"
            f"🛬 <b>العودة:</b> <code>{ret}</code>\n\n"
            f"⚠️ لم يتم العثور على رحلات مباشرة متوفرة في هذين التاريخين، أو أن المقاعد قد نفدت بالكامل.\n\n"
            f"🔗 <a href='{url}'>فحص رحلات الترانزيت أو التواريخ المجاورة على Google Flights ↗</a>"
        )
        send_telegram_msg(msg, high_priority=True)
        return

    available_flights.sort(key=lambda x: x["price"])
    cheapest = available_flights[0]

    lines = [
        f"🎯 <b>نتيجة الفحص المخصص الفوري (القصيم ⇄ جدة)</b>\n",
        f"🛫 <b>تاريخ الذهاب:</b> <code>{dep}</code>\n"
        f"🛬 <b>تاريخ العودة:</b> <code>{ret}</code>\n",
        f"🏆 <b>أفضل خيار متاح:</b>\n"
        f"   💰 <b>{cheapest['price']} ر.س إجمالي</b> — ✈️ {cheapest['airline']} (⏰ {cheapest['time_ar']})\n"
    ]

    seen_airlines = set()
    lines.append("📋 <b>الخيارات المباشرة المتوفرة:</b>")
    for f in available_flights:
        air_name = f["airline"]
        if air_name not in seen_airlines:
            seen_airlines.add(air_name)
            bag_txt = " (شحن 23kg مجاناً 🎁)" if "السعودية" in air_name else f" (+{BAGGAGE_FEES.get(air_name, 140)} حقيبة)"
            dir_link = get_direct_booking_link(air_name, dep, ret)
            lines.append(f"• <b>{f['price']} ر.س</b> — ✈️ {air_name}{bag_txt}\n  🔗 <a href='{dir_link}'>حجز مباشر من {air_name} ↗</a>")

    cal_link = get_google_calendar_link("رحلة مخصصة", cheapest["airline"], dep, ret, cheapest["price"], url)
    lines.append(f"\n📅 <a href='{cal_link}'>إضافة موعد الرحلة لتقويم Google ↗</a>")
    lines.append(f"🔍 <a href='{url}'>مقارنة كافة تفاصيل الرحلة على Google Flights ↗</a>")

    send_telegram_msg("\n".join(lines), high_priority=True)
    print("✅ تم إرسال تقرير الفحص المخصص للتيليجرام بنجاح!")


def run_cloud_scan():
    if CUSTOM_DEP and CUSTOM_RET:
        run_custom_date_probe(CUSTOM_DEP, CUSTOM_RET)
        return

    start_time = time.time()
    now_ts = int(start_time)
    print("☁️ بدء تشغيل الرادار السحابي وصائد القيعان (درع التخفي مفعّل)...")
    pairs = get_all_monitored_pairs(months_ahead=3)
    results = []
    hist_state = load_history()
    price_history = hist_state.get("prices", {})
    alert_history = hist_state.get("alerts", {})

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-infobars"
                ]
            )

            context = create_stealth_context(browser)
            page = context.new_page()

            for idx, (dep, ret, trip_type) in enumerate(pairs, 1):
                url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep}%20through%20{ret}%20nonstop&curr=SAR&hl=en&gl=sa"
                cheapest_flight = None

                # محاولة فحص مع دعم إعادة المحاولة الذكية (Exponential Backoff)
                for attempt in range(2):
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=18000)

                        if "consent.google.com" in page.url:
                            try:
                                page.locator('button:has-text("Accept all"), button:has-text("I agree")').first.click(timeout=3000)
                                page.wait_for_load_state("domcontentloaded", timeout=5000)
                            except Exception:
                                pass

                        # سلوك تفاعلي بشري
                        page.mouse.wheel(0, random.randint(100, 250))

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
                            break  # تم الرصد بنجاح

                    except Exception:
                        if attempt == 0:
                            time.sleep(2)  # انتظار وإعادة المحاولة

                if cheapest_flight:
                    cur_p = cheapest_flight["price"]
                    air = cheapest_flight["airline"]
                    f_time = cheapest_flight["time_ar"]
                    flight_key = f"{dep}_{ret}"
                    print(f"[{idx}/{len(pairs)}] ✅ رصد: {trip_type} -> {cur_p} ر.س ({air})")

                    # ==========================================
                    # 🎯 محرك صائد القيعان والهبوط الساحق والنسب المئوية
                    # ==========================================
                    prev_p = price_history.get(flight_key)
                    last_alert = alert_history.get(flight_key, {})
                    last_alert_p = last_alert.get("price", 9999)
                    last_alert_ts = last_alert.get("time", 0)

                    # حساب نسبة وفرق الهبوط المئوي
                    diff_amount = (prev_p - cur_p) if prev_p else 0
                    drop_percentage = round((diff_amount / prev_p) * 100) if (prev_p and prev_p > 0) else 0

                    # تصنيف مستويات التنبيه
                    is_crash_drop = drop_percentage >= 50       # هبوط ساحق (50% أو أكثر)
                    is_deep_drop = diff_amount >= 60           # هبوط نقدي ملحوظ (60 ر.س فأكثر)
                    is_rock_bottom = cur_p <= 358              # سعر القاع الترويجي المطلق

                    should_alert = False
                    if is_crash_drop or is_deep_drop or is_rock_bottom:
                        if cur_p < last_alert_p or (now_ts - last_alert_ts) >= 86400:
                            should_alert = True

                    if should_alert:
                        direct_url = get_direct_booking_link(air, dep, ret)

                        # صياغة عنوان وطبيعة التنبيه بإنذار مميز حسب قوة الصفقة
                        if is_crash_drop:
                            header_title = "🚨🔥 <b>صفقة الموسم: انهيار سعري ساحق (خطأ تسعيري محتمل)!</b>"
                            reason = f"💥 <b>هبوط جنوني بنسبة {drop_percentage}%</b> (وفّرت <b>{diff_amount} ر.س</b> دفعة واحدة!)"
                        elif is_rock_bottom:
                            header_title = "🔥 <b>صائد القيعان التلقائي: بلوغ سعر القاع التاريخي!</b>"
                            reason = "🎯 <b>السعر وصل للقاع الترويجي الأدنى (358 ر.س أو أقل ذهاب وعودة).</b>"
                        else:
                            header_title = "⚡📉 <b>رادار الهبوط السريع: انخفاض ملحوظ في السعر</b>"
                            reason = f"📉 هبوط بمقدار <b>{diff_amount} ر.س</b> ({drop_percentage}%) عن آخر فحص."

                        flash_msg = (
                            f"{header_title}\n\n"
                            f"🗓 <b>{trip_type}</b>\n"
                            f"🛫 <b>الذهاب:</b> <code>{dep}</code> (⏰ {f_time})\n"
                            f"🛬 <b>العودة:</b> <code>{ret}</code>\n"
                            f"✈️ <b>الناقل:</b> {air}\n"
                            f"💰 <b>السعر الجديد:</b> <b>{cur_p} ر.س فقط!</b> " + (f"<i>(كان {prev_p} ر.س)</i>\n" if prev_p else "\n") +
                            f"📢 <b>طبيعة الصفقة:</b> {reason}\n\n"
                            f"⚡ <i>يُوصى بالحجز فوراً قبل تعديل المقاعد أو إغلاق الفئة.</i>\n\n"
                            f"✈️ <a href='{direct_url}'><b>اقتناص التذكرة فوراً من موقع {air} ↗</b></a>\n"
                            f"🔍 <a href='{url}'>فحص ومقارنة البدائل على Google Flights ↗</a>"
                        )
                        send_telegram_msg(flash_msg, high_priority=True)
                        alert_history[flight_key] = {"price": cur_p, "time": now_ts}

                    price_history[flight_key] = cur_p
                    results.append({
                        "نوع العطلة": trip_type,
                        "تاريخ الذهاب": dep,
                        "وقت الإقلاع": f_time,
                        "تاريخ العودة": ret,
                        "الناقل": air,
                        "السعر": cur_p,
                        "الرابط": url
                    })

                # تأخير زمني بشري عشوائي لمنع رصد التتابع الآلي
                time.sleep(random.uniform(0.7, 1.6))

            browser.close()

    except Exception as fatal_e:
        print(f"❌ خطأ جسيم في تشغيل المتصفح: {fatal_e}")
        send_telegram_msg(
            f"⚠️🚨 <b>تنبيه عطل طارئ في رادار الطيران السحابي</b>\n\n"
            f"فشلت دورة الفحص في خوادم GitHub Actions:\n"
            f"<code>{fatal_e}</code>",
            high_priority=True
        )
        raise fatal_e

    duration = round(time.time() - start_time, 1)
    hist_state["prices"] = price_history
    hist_state["alerts"] = alert_history
    save_history(hist_state)

    if not results:
        send_telegram_msg("⚠️🚨 تحذير: اكتملت دورة الفحص ولكن لم يتم رصد أي رحلة (0 نتائج)!", high_priority=True)
    else:
        results = sorted(results, key=lambda x: x["السعر"])
        sync_to_google_sheets(results, duration)
        df = pd.DataFrame(results)

        ksa_now = datetime.datetime.utcnow() + datetime.timedelta(hours=3)
        is_manual_trigger = os.environ.get("GITHUB_EVENT_NAME") == "workflow_dispatch"
        is_morning_window = (8 <= ksa_now.hour < 10)

        if is_manual_trigger or is_morning_window:
            generate_price_chart(results)
            briefing_text = build_briefing_message(results)
            if os.path.exists(CHART_IMAGE_PATH):
                caption = f"📈 <b>مخطط حركة أسعار (القصيم ⇄ جدة)</b>\n• أديل (أخضر) | السعودية (كحلي) | ناس (ذهبي)"
                send_telegram_photo(CHART_IMAGE_PATH, caption=caption)
            send_telegram_msg(briefing_text)

        df.to_excel(EXCEL_FILE, index=False)
        # تصدير نسخة JSON لخدمة تطبيق التيليجرام المصغر (Telegram Mini App)
        ksa_now = datetime.datetime.utcnow() + datetime.timedelta(hours=3)
        mini_app_payload = {
            "updated_at": ksa_now.strftime("%Y-%m-%d %I:%M %p"),
            "flights": results
        }
        with open("flights_data.json", "w", encoding="utf-8") as f:
            json.dump(mini_app_payload, f, ensure_ascii=False, indent=2)
        print("📱 تم توليد وتحديث ملف بيانات التطبيق المصغر (flights_data.json) بنجاح!")
        print(f"✅ اكتملت الدورة السحابية وصيد القيعان بنجاح في {duration} ثانية! تم رصد {len(results)} رحلة.")


if __name__ == "__main__":
    run_cloud_scan()
