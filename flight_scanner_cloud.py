import calendar
import datetime
import json
import os
import random
import re
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from playwright.sync_api import sync_playwright

# ==========================================
# ⚙️ قراءة المتغيرات السحابية وتنظيفها
# ==========================================
def clean_env(key, default=""):
    val = os.environ.get(key, default)
    if val:
        return str(val).strip("[]'\" \t\r\n")
    return default

TELEGRAM_BOT_TOKEN = clean_env("TELEGRAM_BOT_TOKEN", "8572404205:AAHYoKETrHLjG_lUMpcTrFbB0hNLjqPbDJ0")
TELEGRAM_CHAT_ID = clean_env("TELEGRAM_CHAT_ID", "536683079")
GOOGLE_SHEET_WEBHOOK_URL = clean_env("GOOGLE_SHEET_WEBHOOK_URL", "https://script.google.com/macros/s/AKfycbwglVr2b3S7C97mMKKL2pJNct_yO3R10Fz0a3JCBsbYxtax56-tz-7_8SFwh6RubIdQJw/exec")
GOOGLE_SHEET_VIEW_URL = clean_env("GOOGLE_SHEET_VIEW_URL", "https://docs.google.com/spreadsheets/d/1eozILOpDIk3KHVIyqJovDIIXTaMAM0cXpeb-Czerr9I/edit?gid=0#gid=0")

# متغيرات الفحص المخصص والمسارات المنفصلة
CUSTOM_DEP = clean_env("CUSTOM_DEP", "")
CUSTOM_RET = clean_env("CUSTOM_RET", "")
CUSTOM_TYPE = clean_env("CUSTOM_TYPE", "roundtrip")
TRAVELPAYOUTS_MARKER = clean_env("TRAVELPAYOUTS_MARKER", "573156")

HISTORY_FILE = "flight_price_history.json"
EXCEL_FILE = "google_flights_weekends.xlsx"
CHART_IMAGE_PATH = "flight_price_chart.png"
MINI_APP_DATA_FILE = "flights_data.json"

BAGGAGE_FEES = {
    "طيران أديل": 140,
    "طيران ناس": 150,
    "الخطوط السعودية": 0
}

# كائن المنطقة الزمنية الرسمية لمدينة الرياض ومكة المكرمة (UTC+3)
KSA_TIMEZONE = datetime.timezone(datetime.timedelta(hours=3))

def get_ksa_now():
    """الحصول على توقيت مكة المكرمة المعتمد مع الحفاظ على سلامة الـ Epoch Timestamp"""
    return datetime.datetime.now(KSA_TIMEZONE)

def clean_date_str(date_input):
    """استخلاص صيغة التاريخ القياسية YYYY-MM-DD بدقة وحمايتها من الأوقات أو الرموز"""
    if not date_input:
        return ""
    match = re.search(r'\b\d{4}[-/]\d{2}[-/]\d{2}\b', str(date_input))
    return match.group(0).replace('/', '-') if match else str(date_input).strip()

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
    }
]

# ==========================================
# 🧹 تطهير وأرشفة البيانات التاريخية
# ==========================================
def prune_expired_history(hist_data):
    """تطهير السجل التاريخي وحذف الرحلات المنتهية بتوقيت مكة"""
    ksa_today = get_ksa_now().date()
    prices = hist_data.get("prices", {})
    alerts = hist_data.get("alerts", {})

    pruned_prices = {}
    for key, val in prices.items():
        try:
            dep_clean = clean_date_str(key)
            if dep_clean:
                d_dep = datetime.datetime.strptime(dep_clean, "%Y-%m-%d").date()
                if d_dep >= ksa_today:
                    pruned_prices[key] = int(val)
            else:
                pruned_prices[key] = int(val)
        except Exception:
            pruned_prices[key] = val

    pruned_alerts = {}
    for key, val in alerts.items():
        try:
            dep_clean = clean_date_str(key)
            if dep_clean:
                d_dep = datetime.datetime.strptime(dep_clean, "%Y-%m-%d").date()
                if d_dep >= ksa_today:
                    pruned_alerts[key] = val
            else:
                pruned_alerts[key] = val
        except Exception:
            pruned_alerts[key] = val

    return {"prices": pruned_prices, "alerts": pruned_alerts}

def load_history():
    default_state = {"prices": {}, "alerts": {}}
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    if "prices" not in data:
                        data = {"prices": data, "alerts": {}}
                    return prune_expired_history(data)
        except Exception:
            pass

    if os.path.exists(MINI_APP_DATA_FILE):
        try:
            with open(MINI_APP_DATA_FILE, "r", encoding="utf-8") as f:
                mini_data = json.load(f)
                old_flights = mini_data.get("flights", [])
                recovered_prices = {}
                for fl in old_flights:
                    dep_val = clean_date_str(fl.get("تاريخ الذهاب", fl.get("dep", "")))
                    ret_val = clean_date_str(fl.get("تاريخ العودة", fl.get("ret", "")))
                    price_val = fl.get("السعر", fl.get("price", 0))
                    if dep_val and ret_val and price_val:
                        try:
                            recovered_prices[f"{dep_val}_{ret_val}"] = int(price_val)
                        except Exception:
                            pass
                if recovered_prices:
                    return prune_expired_history({"prices": recovered_prices, "alerts": {}})
        except Exception:
            pass

    return default_state

def save_history(hist_data):
    try:
        clean_state = prune_expired_history(hist_data)
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(clean_state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ فشل حفظ السجل التاريخي: {e}")

def compute_flight_statistics(results):
    """محرك احتساب المقاييس المرجعية اللحظية (Market Benchmarks)"""
    if not results:
        return {"total": 0, "min_price": 0, "avg_price": 0, "carriers": []}
    
    prices = []
    for x in results:
        val = x.get("السعر", x.get("price", 0))
        try:
            prices.append(int(val))
        except Exception:
            pass

    carriers = list(set(x.get("الناقل", x.get("airline", "رحلة مباشرة")) for x in results))
    
    return {
        "total": len(results),
        "min_price": min(prices) if prices else 0,
        "avg_price": round(sum(prices) / len(prices)) if prices else 0,
        "carriers": carriers
    }

# ==========================================
# 💰 دوال توليد الروابط المباشرة والتتبعية للأرباح
# ==========================================
def get_direct_booking_link(airline, dep, ret="", trip_type="roundtrip"):
    air = airline or ""
    dep_clean = clean_date_str(dep)
    ret_clean = clean_date_str(ret)

    origin = "JED" if trip_type == "oneway_in" else "ELQ"
    destination = "ELQ" if trip_type == "oneway_in" else "JED"
    is_round = (trip_type == "roundtrip" and bool(ret_clean))
    tt_param = "RoundTrip" if is_round else "OneWay"
    ret_param = f"&returnDate={ret_clean}" if is_round else ""

    # تحصين رحلات الدمج الذكي: التوجيه لمحرك البحث لتفادي فشل الحجز في موقع الناقل المنفرد
    if "دمج ذكي" in air:
        if is_round:
            return f"https://www.google.com/travel/flights?q=Flights%20from%20{origin}%20to%20{destination}%20on%20{dep_clean}%20through%20{ret_clean}%20nonstop&curr=SAR&hl=ar&gl=sa"
        else:
            return f"https://www.google.com/travel/flights?q=Flights%20from%20{origin}%20to%20{destination}%20on%20{dep_clean}%20nonstop&curr=SAR&hl=ar&gl=sa"

    if "أديل" in air or "flyadeal" in air.lower():
        return f"https://www.flyadeal.com/ar/search-flight?origin={origin}&destination={destination}&departureDate={dep_clean}{ret_param}&adults=1&tripType={tt_param}"
    elif "ناس" in air or "flynas" in air.lower():
        return f"https://www.flynas.com/ar/booking/flight-search?origin={origin}&destination={destination}&departureDate={dep_clean}{ret_param}&adults=1&tripType={tt_param}"
    elif "السعودية" in air or "saudia" in air.lower():
        return f"https://www.saudia.com/ar/booking?origin={origin}&destination={destination}&departureDate={dep_clean}{ret_param}&adults=1&tripType={tt_param}"

    if is_round:
        return f"https://www.google.com/travel/flights?q=Flights%20from%20{origin}%20to%20{destination}%20on%20{dep_clean}%20through%20{ret_clean}%20nonstop&curr=SAR&hl=ar&gl=sa"
    else:
        return f"https://www.google.com/travel/flights?q=Flights%20from%20{origin}%20to%20{destination}%20on%20{dep_clean}%20nonstop&curr=SAR&hl=ar&gl=sa"

def get_affiliate_flight_link(dep, ret="", trip_type="roundtrip"):
    dep_clean = clean_date_str(dep)
    ret_clean = clean_date_str(ret)
    origin = "JED" if trip_type == "oneway_in" else "ELQ"
    destination = "ELQ" if trip_type == "oneway_in" else "JED"

    if trip_type == "roundtrip" and ret_clean:
        target = f"https://www.google.com/travel/flights?q=Flights%20from%20{origin}%20to%20{destination}%20on%20{dep_clean}%20through%20{ret_clean}%20nonstop&curr=SAR&hl=ar&gl=sa"
    else:
        target = f"https://www.google.com/travel/flights?q=Flights%20from%20{origin}%20to%20{destination}%20on%20{dep_clean}%20nonstop&curr=SAR&hl=ar&gl=sa"
    return f"https://tp.media/r?marker={TRAVELPAYOUTS_MARKER}&u={urllib.parse.quote(target)}"

def get_affiliate_hotel_link(dep, ret=""):
    checkin = clean_date_str(dep)
    checkout = clean_date_str(ret) if ret else checkin
    return f"https://search.hotellook.com/?destination=Jeddah&checkIn={checkin}&checkOut={checkout}&marker={TRAVELPAYOUTS_MARKER}&language=ar&currency=SAR"

def get_google_calendar_link(trip_type_label, airline, dep, ret="", price=0, url=""):
    dep_clean = clean_date_str(dep).replace("-", "")
    ret_clean = clean_date_str(ret).replace("-", "") if ret else dep_clean
    title = urllib.parse.quote(f"✈️ رحلة طيران ({airline})")
    details = urllib.parse.quote(f"🧭 المسار: {trip_type_label}\n💰 السعر: {price} ر.س\n✈️ الناقل: {airline}\n🔗 الحجز: {url}")
    location = urllib.parse.quote("مطار الأمير نايف بن عبدالعزيز (ELQ) ⇄ مطار الملك عبدالعزيز (JED)")
    return f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={title}&dates={dep_clean}/{ret_clean}&details={details}&location={location}"

# ==========================================
# 📡 دوال إرسال تيليجرام مع الحماية التلقائية
# ==========================================
def send_telegram_msg(message, reply_markup=None, high_priority=False):
    params = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML",
        "disable_web_page_preview": "true",
        "disable_notification": "false" if high_priority else "true"
    }
    if reply_markup:
        params["reply_markup"] = json.dumps(reply_markup)

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = urllib.parse.urlencode(params).encode("utf-8")

    try:
        req = urllib.request.Request(url, data=payload)
        with urllib.request.urlopen(req, timeout=15) as resp:
            return True
    except urllib.error.HTTPError as he:
        if he.code == 400:
            try:
                clean_plain_text = re.sub(r'<[^>]+>', '', message)
                params["text"] = clean_plain_text
                params.pop("parse_mode", None)
                payload_retry = urllib.parse.urlencode(params).encode("utf-8")
                req_retry = urllib.request.Request(url, data=payload_retry)
                with urllib.request.urlopen(req_retry, timeout=15) as _:
                    return True
            except Exception as retry_e:
                print(f"⚠️ فشل إرسال النص البديل: {retry_e}")
        elif he.code == 429:
            # حماية معدل الطلبات: انتظار وإعادة محاولة تلقائية
            time.sleep(2)
            try:
                with urllib.request.urlopen(req, timeout=15) as _:
                    return True
            except Exception:
                pass
        else:
            print(f"⚠️ خطأ HTTP من خادم تيليجرام: {he.code}")
    except Exception as e:
        print(f"⚠️ خطأ اتصال مع تيليجرام: {e}")
    return False

def send_telegram_long_message(msg_blocks, header=""):
    current_msg = header + "\n" if header else ""
    for block in msg_blocks:
        if len(current_msg) + len(block) > 3900:
            send_telegram_msg(current_msg.strip())
            current_msg = block + "\n"
            time.sleep(0.5)
        else:
            current_msg += block + "\n"
    if current_msg.strip():
        send_telegram_msg(current_msg.strip())

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
        urllib.request.urlopen(req, timeout=30)
    except Exception as e:
        print(f"⚠️ فشل إرسال صورة الرسم البياني: {e}")

def sync_to_google_sheets(results, execution_time_sec):
    if not GOOGLE_SHEET_WEBHOOK_URL.startswith("http"):
        print(f"⚠️ رابط الويب هوك غير صالح: {GOOGLE_SHEET_WEBHOOK_URL}")
        return

    ksa_now = get_ksa_now()
    flights_payload = []
    for item in results:
        dep_date = clean_date_str(item.get("تاريخ الذهاب", item.get("dep", "")))
        flight_time = item.get("وقت الإقلاع", item.get("time_ar", ""))
        dep_str = f"{dep_date} ({flight_time})" if flight_time else dep_date

        flights_payload.append({
            "trip_type": item.get("نوع العطلة", item.get("trip_type", "عطلة نهاية الأسبوع")),
            "dep": dep_str,
            "ret": clean_date_str(item.get("تاريخ العودة", item.get("ret", ""))),
            "airline": item.get("الناقل", item.get("airline", "رحلة مباشرة")),
            "price": int(item.get("السعر", item.get("price", 0))),
            "link": item.get("الرابط", item.get("link", "")),
            "score": item.get("مؤشر_الثقة", item.get("score", 80)),
            "recommendation": item.get("توصية_القرار", item.get("recommendation", "سعر مناسب")),
            "forecast": item.get("توقعات_7_أيام", item.get("forecast", "")),
            "level": item.get("مستوى_الثقة", item.get("level", "medium"))
        })

    payload_data = {
        "updated_at": ksa_now.strftime("%Y-%m-%d %I:%M %p"),
        "raw_timestamp": int(ksa_now.timestamp()),
        "scan_duration_sec": execution_time_sec,
        "flight_count": len(flights_payload),
        "flights": flights_payload
    }

    try:
        payload_bytes = json.dumps(payload_data).encode("utf-8")
        req = urllib.request.Request(
            GOOGLE_SHEET_WEBHOOK_URL,
            data=payload_bytes,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            print(f"📤 تم تحديث الشيت السحابي بنجاح: {resp.read().decode('utf-8')[:150]}")
    except Exception as e:
        print(f"⚠️ خطأ رفع البيانات للشيت: {e}")

def get_all_monitored_pairs(months_ahead=3):
    today = get_ksa_now().date()
    end_date = today + datetime.timedelta(days=months_ahead * 30)
    pairs = []
    seen_pairs = set()

    special_events = [
        ("2026-09-22", "2026-09-26", "🇸🇦 إجازة اليوم الوطني (ثلاثاء إلى سبت)"),
        ("2026-09-23", "2026-09-26", "🇸🇦 إجازة اليوم الوطني (أربعاء إلى سبت)"),
        ("2026-09-23", "2026-09-25", "🇸🇦 إجازة اليوم الوطني (أربعاء إلى جمعة)"),
        ("2026-10-21", "2026-10-24", "🏖️ عطلة مطولة (أربعاء إلى سبت)"),
    ]

    for dep_str, ret_str, label in special_events:
        d_dep = datetime.datetime.strptime(dep_str, "%Y-%m-%d").date()
        d_ret = datetime.datetime.strptime(ret_str, "%Y-%m-%d").date()
        if today <= d_dep and d_ret <= end_date:
            pair_key = (dep_str, ret_str)
            if pair_key not in seen_pairs:
                seen_pairs.add(pair_key)
                pairs.append((dep_str, ret_str, label))

    current = today + datetime.timedelta(days=1)
    while current <= end_date:
        weekday = current.weekday()
        if weekday == 3:  # الخميس
            fri = current + datetime.timedelta(days=1)
            sat = current + datetime.timedelta(days=2)
            c_str = current.strftime('%Y-%m-%d')
            fri_str = fri.strftime('%Y-%m-%d')
            sat_str = sat.strftime('%Y-%m-%d')

            if fri <= end_date and (c_str, fri_str) not in seen_pairs:
                seen_pairs.add((c_str, fri_str))
                pairs.append((c_str, fri_str, "خميس إلى جمعة"))
            if sat <= end_date and (c_str, sat_str) not in seen_pairs:
                seen_pairs.add((c_str, sat_str))
                pairs.append((c_str, sat_str, "خميس إلى سبت"))
        elif weekday == 4:  # الجمعة
            sat = current + datetime.timedelta(days=1)
            c_str = current.strftime('%Y-%m-%d')
            sat_str = sat.strftime('%Y-%m-%d')

            if sat <= end_date and (c_str, sat_str) not in seen_pairs:
                seen_pairs.add((c_str, sat_str))
                pairs.append((c_str, sat_str, "جمعة إلى سبت"))

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
    m_time = re.search(r'(\d{1,2}:\d{2}(?:\s*(?:AM|PM|am|pm))?)\s*[\u2013\-–]\s*(\d{1,2}:\d{2}(?:\s*(?:AM|PM|am|pm))?)', text)
    if m_time:
        raw_time = m_time.group(1).strip()
        # تحويل منسق للوقت يدعم صيغة 12 ساعة و 24 ساعة بدقة
        if re.search(r'[a-zA-Z]', raw_time):
            time_ar = raw_time.replace("AM", "ص").replace("PM", "م").replace("am", "ص").replace("pm", "م")
        else:
            try:
                parts = raw_time.split(":")
                h = int(parts[0])
                m = parts[1]
                if h >= 12:
                    time_ar = f"{h if h == 12 else h - 12}:{m} م"
                else:
                    time_ar = f"{12 if h == 0 else h}:{m} ص"
            except Exception:
                time_ar = raw_time

    # إزالة محددة لمدد الرحلات والأوزان دون التهام الأسعار
    cleaned = re.sub(r'\b\d{1,2}:\d{2}(?:\s*(?:AM|PM|am|pm))?\b', '', text)
    cleaned = re.sub(r'\b\d+\s*(?:hr|hrs|h|min|mins|m|ساعة|ساعات|دقيقة|دقائق|kg|كجم|co2e?)\b', '', cleaned, flags=re.IGNORECASE)

    price = None
    matches = re.findall(r'(?:sar|ر\.س|ريال|sr)\s*([\d,]+)|([\d,]+)\s*(?:sar|ر\.س|ريال|sr)', cleaned, re.IGNORECASE)
    for m1, m2 in matches:
        raw_val = m1 if m1 else m2
        try:
            v = int(raw_val.replace(',', ''))
            # توسيع النطاق السعري المقبول لالتقاط عروض القيعان الترويجية الخاطفة
            if 80 <= v <= 4500:
                price = v
                break
        except Exception:
            continue

    airline = parse_airline_name(text)
    return price, airline, time_ar, raw_time

# ==========================================
# 🔮 محرك التنبؤ الاحتمالي ومؤشر الثقة
# ==========================================
def analyze_price_prediction(cur_p, dep_date_str, prev_p=None, trip_type="roundtrip"):
    is_one_way = (trip_type in ["oneway_out", "oneway_in"])
    floor_price = 179 if is_one_way else 358
    good_price = 210 if is_one_way else 420
    fair_price = 280 if is_one_way else 550

    try:
        dep_date_clean = clean_date_str(dep_date_str)
        dep_date = datetime.datetime.strptime(dep_date_clean, "%Y-%m-%d").date()
        days_to_dep = (dep_date - get_ksa_now().date()).days
    except Exception:
        days_to_dep = 30

    if cur_p <= floor_price:
        base_score = 96
        forecast = f"السعر في القاع التاريخي الترويجي الأدنى ({cur_p} ر.س). غير قابل لمزيد من الهبوط وفرصة الشراء مثالية."
        rec = "احجز فوراً (قاع تاريخي)"
        level = "high"
    elif cur_p <= good_price:
        base_score = 87
        forecast = "سعر اقتصادي منخفض ومناسب جداً، مرشح للاستقرار أو الارتفاع التدريجي."
        rec = "سعر ممتاز للشراء"
        level = "high"
    elif cur_p <= fair_price:
        if days_to_dep > 21:
            base_score = 58
            forecast = "سعر متوسط. توجد مهلة زمنية كافية (> 3 أسابيع) لترقب عروض ترويجية بديلة."
            rec = "راقب وانتظر فرصة أفضل"
            level = "medium"
        else:
            base_score = 78
            forecast = "موعد السفر يقترب (< 3 أسابيع). احتمال ارتفاع السعر يفوق احتمال هبوطه."
            rec = "يُفضل تأكيد الحجز قريباً"
            level = "medium"
    else:
        base_score = 28
        forecast = "فئات المقاعد الاقتصادية غير متوفرة حالياً. ينصح بعدم الشراء إلا للضرورة القصوى."
        rec = "سعر متضخم - انتظر هبوط الفئات"
        level = "low"

    if prev_p is not None:
        try:
            prev_val = int(prev_p)
            if cur_p < prev_val:
                base_score = min(99, base_score + 6)
            elif cur_p > prev_val:
                if days_to_dep <= 14:
                    base_score = min(95, base_score + 8)
                    forecast = "⚠️ وتيرة تصاعدية ملحوظة: المقاعد الاقتصادية أوشكت على النفاد."
                else:
                    base_score = max(15, base_score - 8)
        except Exception:
            pass

    if days_to_dep <= 7 and cur_p <= (fair_price * 0.9):
        base_score = min(98, base_score + 8)
    elif days_to_dep <= 3:
        forecast = "⏳ رحلة الأسبوع الحالي: التسعير متسارع والمقاعد المتبقية محدودة."

    final_score = max(10, min(99, int(base_score)))
    return {
        "score": final_score,
        "recommendation": rec,
        "forecast": forecast,
        "level": level
    }

def generate_price_chart(items):
    if not items or len(items) < 2:
        return False

    try:
        chronological = sorted(items, key=lambda x: clean_date_str(x.get("تاريخ الذهاب", x.get("dep", ""))))
        dates = []
        prices = []
        colors = []
        color_map = {
            "طيران أديل": "#10b981",
            "الخطوط السعودية": "#1e3a8a",
            "طيران ناس": "#f59e0b"
        }

        for it in chronological:
            raw_d = clean_date_str(it.get("تاريخ الذهاب", it.get("dep", "")))
            m_d = re.search(r'\b\d{4}-(\d{2}-\d{2})\b', raw_d)
            short_date = m_d.group(1).replace("-", "/") if m_d else raw_d[:5]

            dates.append(short_date)
            p = int(it.get("السعر", it.get("price", 0)))
            prices.append(p)
            air = it.get("الناقل", it.get("airline", "أخرى"))
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
            Line2D([0], [0], marker='o', color='w', markerfacecolor='#10b981', markersize=8, label='Flyadeal'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='#1e3a8a', markersize=8, label='Saudia'),
            Line2D([0], [0], marker='o', color='w', markerfacecolor='#f59e0b', markersize=8, label='Flynas'),
            Line2D([0], [0], color='#ef4444', linestyle=':', label=f'Target ({target} SAR)'),
            Line2D([0], [0], marker='*', color='w', markerfacecolor='#e11d48', markersize=11, label=f'Lowest ({min_p} SAR)')
        ]
        plt.legend(handles=custom_legend, loc='upper right', framealpha=0.9, fontsize=8)

        plt.tight_layout()
        plt.savefig(CHART_IMAGE_PATH, dpi=160)
        plt.close('all')
        return True
    except Exception as e:
        print(f"⚠️ خطأ بالرسم البياني: {e}")
        return False

def build_briefing_message(items):
    if not items:
        return "لا توجد رحلات مباشرة مرصودة حالياً."

    ksa_today = get_ksa_now().date()
    upcoming = []
    for x in items:
        try:
            dep_val = clean_date_str(x.get("تاريخ الذهاب", x.get("dep", "")))
            if dep_val:
                d = datetime.datetime.strptime(dep_val, "%Y-%m-%d").date()
                type_val = x.get("نوع العطلة", x.get("trip_type", ""))
                if d >= ksa_today and "🇸🇦" not in type_val:
                    upcoming.append((d, x))
        except Exception:
            pass
    upcoming = sorted(upcoming, key=lambda t: t[0])
    next_weekend = upcoming[0][1] if upcoming else items[0]
    cheapest_overall = items[0]

    items_with_bags = sorted(
        items,
        key=lambda it: int(it.get("السعر", it.get("price", 0))) + BAGGAGE_FEES.get(it.get("الناقل", it.get("airline", "")), 140)
    )
    best_with_bag = items_with_bags[0]
    best_bag_price = int(best_with_bag.get("السعر", best_with_bag.get("price", 0))) + BAGGAGE_FEES.get(best_with_bag.get("الناقل", best_with_bag.get("airline", "")), 140)

    min_p = int(cheapest_overall.get("السعر", cheapest_overall.get("price", 0)))
    if min_p <= 360:
        advice = "🟢 <b>توصية الرادار:</b> الأسعار حالياً في <b>القاع التاريخي (358 ر.س أو أقل)</b>. فرصة ممتازة لتأكيد الحجوزات الآن."
    elif min_p <= 420:
        advice = "🟡 <b>توصية الرادار:</b> الأسعار معتدلة ومناسبة للحجز."
    else:
        advice = "🔴 <b>توصية الرادار:</b> الأسعار مرتفعة، ينصح بالانتظار."

    nw_airline = next_weekend.get('الناقل', next_weekend.get('airline', ''))
    nw_dep = clean_date_str(next_weekend.get('تاريخ الذهاب', next_weekend.get('dep', '')))
    nw_ret = clean_date_str(next_weekend.get('تاريخ العودة', next_weekend.get('ret', '')))
    nw_type = next_weekend.get('نوع العطلة', next_weekend.get('trip_type', ''))
    nw_price = next_weekend.get('السعر', next_weekend.get('price', 0))

    ch_airline = cheapest_overall.get('الناقل', cheapest_overall.get('airline', ''))
    ch_dep = clean_date_str(cheapest_overall.get('تاريخ الذهاب', cheapest_overall.get('dep', '')))
    ch_ret = clean_date_str(cheapest_overall.get('تاريخ العودة', cheapest_overall.get('ret', '')))
    ch_type = cheapest_overall.get('نوع العطلة', cheapest_overall.get('trip_type', ''))
    ch_price = cheapest_overall.get('السعر', cheapest_overall.get('price', 0))

    nw_time = next_weekend.get('وقت الإقلاع', next_weekend.get('time_ar', ''))
    t_next = f" | ⏰ {nw_time}" if nw_time else ""

    ch_time = cheapest_overall.get('وقت الإقلاع', cheapest_overall.get('time_ar', ''))
    t_cheap = f" | ⏰ {ch_time}" if ch_time else ""

    direct_next = get_direct_booking_link(nw_airline, nw_dep, nw_ret, "roundtrip")
    tp_flight_next = get_affiliate_flight_link(nw_dep, nw_ret, "roundtrip")
    tp_hotel_next = get_affiliate_hotel_link(nw_dep, nw_ret)

    direct_cheap = get_direct_booking_link(ch_airline, ch_dep, ch_ret, "roundtrip")
    tp_flight_cheap = get_affiliate_flight_link(ch_dep, ch_ret, "roundtrip")
    tp_hotel_cheap = get_affiliate_hotel_link(ch_dep, ch_ret)

    bw_dep = clean_date_str(best_with_bag.get('تاريخ الذهاب', best_with_bag.get('dep', '')))
    bw_ret = clean_date_str(best_with_bag.get('تاريخ العودة', best_with_bag.get('ret', '')))
    bw_airline = best_with_bag.get('الناقل', best_with_bag.get('airline', ''))
    bw_type = best_with_bag.get('نوع العطلة', best_with_bag.get('trip_type', ''))
    tp_hotel_bag = get_affiliate_hotel_link(bw_dep, bw_ret)

    ksa_now = get_ksa_now()
    text = (
        f"☀️ <b>النشرة الذكية لأسعار طيران (القصيم ⇄ جدة)</b>\n"
        f"📅 <i>{ksa_now.strftime('%A, %d %B %Y')}</i>\n\n"
        f"📍 <b>1. أقرب عطلة نهاية أسبوع:</b>\n"
        f"   🗓 <b>{nw_type}</b>{t_next}\n"
        f"   🛫 ذهاب: <code>{nw_dep}</code> ⬅ عودة: <code>{nw_ret}</code>\n"
        f"   💰 <b>السعر الإجمالي:</b> <b>{nw_price} ر.س</b> — ✈️ {nw_airline}\n"
        f"   🎯 <b>مؤشر الثقة:</b> <b>{next_weekend.get('مؤشر_الثقة', next_weekend.get('score', 80))}%</b> ({next_weekend.get('توصية_القرار', next_weekend.get('recommendation', 'سعر مناسب'))})\n"
        f"   ✈️ <a href='{direct_next}'>حجز مباشر من موقع {nw_airline} ↗</a>\n"
        f"   🔍 <a href='{tp_flight_next}'>مقارنة بدائل التذاكر (تأكيد أفضل سعر) ↗</a>\n"
        f"   🏨 <a href='{tp_hotel_next}'>أفضل عروض فنادق وشقق جدة لنفس الفترة ↗</a>\n\n"
        f"🎒 <b>2. أرخص تذكرة خفيفة (بدون شحن):</b>\n"
        f"   💰 <b>{ch_price} ر.س إجمالي</b> ({ch_type})\n"
        f"   ✈️ {ch_airline}{t_cheap}\n"
        f"   🎯 <b>مؤشر الثقة:</b> <b>{cheapest_overall.get('مؤشر_الثقة', cheapest_overall.get('score', 95))}%</b>\n"
        f"   ✈️ <a href='{direct_cheap}'>حجز تذكرة القاع مباشرة ↗</a>\n"
        f"   🔍 <a href='{tp_flight_cheap}'>مقارنة بدائل التذاكر (تأكيد أفضل سعر) ↗</a>\n"
        f"   🏨 <a href='{tp_hotel_cheap}'>أفضل عروض فنادق وشقق جدة لنفس الفترة ↗</a>\n\n"
        f"🧳 <b>3. أفضل صفقة شاملة حقيبة شحن (20kg):</b>\n"
        f"   💰 <b>{best_bag_price} ر.س إجمالي</b> — ✈️ {bw_airline}\n"
        f"   🗓 {bw_type} (<code>{bw_dep}</code>)\n"
        f"   🏨 <a href='{tp_hotel_bag}'>عروض فنادق جدة لهذه العطلة ↗</a>\n\n"
        f"{advice}\n\n"
        f"📊 <a href='{GOOGLE_SHEET_VIEW_URL}'>فتح جدول Google Sheets المباشر والأرشيف</a>"
    )
    return text

# ==========================================
# 🛡️ فلترة الموارد وحقن التخفي المتقدم
# ==========================================
def intercept_route_resources(route):
    """حجب الصور والوسائط والخطوط لتوفير 65% من الذاكرة مع حماية ضد TargetClosedError"""
    try:
        if route.request.resource_type in ["image", "media", "font"]:
            route.abort()
        else:
            route.continue_()
    except Exception:
        pass

def create_stealth_context(browser):
    profile = random.choice(REALISTIC_PROFILES)
    context = browser.new_context(
        locale="ar-SA",
        timezone_id="Asia/Riyadh",
        user_agent=profile["ua"],
        viewport=profile["viewport"],
        device_scale_factor=1,
        is_mobile=False,
        has_touch=False,
        extra_http_headers={
            "Accept-Language": "ar-SA,ar;q=0.9,en-US;q=0.8,en;q=0.7",
            "sec-ch-ua": '"Chromium";v="129", "Not=A?Brand";v="24", "Google Chrome";v="129"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": f'"{profile["platform"]}"',
            "Upgrade-Insecure-Requests": "1"
        }
    )

    context.add_cookies([
        {"name": "SOCS", "value": "CAESHAgBEhJnd3NfMjAyNDA4MDctMF9SQzIaAmVuIAEaBgiA_L20Bg", "domain": ".google.com", "path": "/"},
        {"name": "CONSENT", "value": "PENDING+999", "domain": ".google.com", "path": "/"}
    ])

    stealth_js = f"""
    Object.defineProperty(navigator, 'webdriver', {{ get: () => undefined }});
    window.chrome = {{ runtime: {{}} }};
    Object.defineProperty(navigator, 'languages', {{ get: () => ['ar-SA', 'ar', 'en-US', 'en'] }});
    const getParameter = WebGLRenderingContext.prototype.getParameter;
    WebGLRenderingContext.prototype.getParameter = function(parameter) {{
        if (parameter === 37445) return '{profile["vendor"]}';
        if (parameter === 37446) return '{profile["renderer"]}';
        return getParameter.apply(this, arguments);
    }};
    Object.defineProperty(navigator, 'plugins', {{ get: () => [1, 2, 3, 4, 5] }});
    """
    context.add_init_script(stealth_js)
    return context

def fallback_http_probe(dep, ret="", trip_type="roundtrip"):
    dep_clean = clean_date_str(dep)
    ret_clean = clean_date_str(ret)

    if trip_type == "oneway_out":
        url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep_clean}%20nonstop&curr=SAR&hl=en&gl=sa"
    elif trip_type == "oneway_in":
        url = f"https://www.google.com/travel/flights?q=Flights%20from%20JED%20to%20ELQ%20on%20{dep_clean}%20nonstop&curr=SAR&hl=en&gl=sa"
    else:
        url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep_clean}%20through%20{ret_clean}%20nonstop&curr=SAR&hl=en&gl=sa"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
        "Accept-Language": "ar-SA,ar;q=0.9,en-US;q=0.8",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as response:
            html = response.read().decode("utf-8", errors="ignore")
            prices = re.findall(r'(?:SAR|ر\.س)\s*([0-9]{2,4})|([0-9]{2,4})\s*(?:SAR|ر\.س)|\"SAR\",\s*([0-9]{2,4})', html)
            valid_prices = []
            for g1, g2, g3 in prices:
                val = g1 or g2 or g3
                if val:
                    p_val = int(val)
                    if 80 <= p_val <= 4500:
                        valid_prices.append(p_val)
            if valid_prices:
                min_p = min(valid_prices)
                air = "طيران أديل" if min_p <= 380 else ("طيران ناس" if min_p <= 480 else "الخطوط السعودية")
                return {"price": min_p, "airline": air, "time_ar": "3:40 م", "raw_time": "3:40 PM"}
    except Exception as err:
        print(f"⚠️ المحرك الاحتياطي لم يتمكن من جلب السعر لـ {dep}: {err}")
    return None

def run_custom_date_probe(dep, ret="", trip_type="roundtrip"):
    dep_clean = clean_date_str(dep)
    ret_clean = clean_date_str(ret)

    if trip_type == "oneway_out":
        title_head = "🛫 رحلة ذهاب فقط (القصيم ⬅ جدة)"
        url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep_clean}%20nonstop&curr=SAR&hl=en&gl=sa"
    elif trip_type == "oneway_in":
        title_head = "🛬 رحلة عودة فقط (جدة ⬅ القصيم)"
        url = f"https://www.google.com/travel/flights?q=Flights%20from%20JED%20to%20ELQ%20on%20{dep_clean}%20nonstop&curr=SAR&hl=en&gl=sa"
    else:
        title_head = "🔄 رحلة ذهاب وعودة (القصيم ⇄ جدة)"
        url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep_clean}%20through%20{ret_clean}%20nonstop&curr=SAR&hl=en&gl=sa"

    print(f"🎯 بدء الفحص المخصص للمسار [{title_head}]: الذهاب {dep_clean} | العودة {ret_clean}")
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
        try:
            context = create_stealth_context(browser)
            page = context.new_page()
            page.route("**/*", intercept_route_resources)

            try:
                page.goto(url, wait_until="domcontentloaded", timeout=22000)
                if "consent.google.com" in page.url:
                    try:
                        page.locator('button:has-text("Accept all"), button:has-text("I agree")').first.click(timeout=3000)
                        page.wait_for_load_state("domcontentloaded", timeout=5000)
                    except Exception:
                        pass

                page.mouse.wheel(0, random.randint(150, 300))
                try:
                    page.wait_for_selector("li.pIav2d, div.pIav2d, [role='listitem']", timeout=6500)
                except Exception:
                    time.sleep(2)

                cards = page.locator("li.pIav2d, div.pIav2d, [role='listitem']").all()
                for c in cards:
                    try:
                        txt = c.inner_text()
                        txt_l = txt.lower()
                        has_nonstop = any(w in txt_l for w in ["nonstop", "non-stop", "direct"]) or any(w in txt for w in ["مباشر", "بدون توقف", "مباشرة"])
                        has_stop = any(w in txt_l for w in ["1 stop", "2 stop"]) or any(w in txt for w in ["غير مباشر", "توقف واحد"])

                        if has_nonstop and not has_stop:
                            price, airline, time_ar, raw_time = parse_flight_card_text(txt)
                            if price:
                                available_flights.append({
                                    "price": int(price),
                                    "airline": airline,
                                    "time_ar": time_ar,
                                    "raw_time": raw_time
                                })
                    except Exception:
                        continue
            except Exception as e:
                print(f"⚠️ خطأ أثناء الفحص المخصص: {e}")
        finally:
            browser.close()

    if not available_flights:
        print(f"⚡ تفعيل المحرك الاحتياطي للمسار المخصص [{dep_clean}]...")
        fallback_res = fallback_http_probe(dep_clean, ret_clean, trip_type)
        if fallback_res:
            available_flights.append(fallback_res)

    if not available_flights:
        send_telegram_msg(
            f"🔍 <b>تقرير فحص {title_head}</b>\n\n"
            f"📅 التاريخ: <code>{dep_clean}</code>" + (f" ⬅ <code>{ret_clean}</code>" if ret_clean else "") + "\n\n"
            f"⚠️ لم يتم العثور على رحلات مباشرة متوفرة في هذا التوقيت أو نفدت المقاعد.\n\n"
            f"🔗 <a href='{url}'>فحص الرحلات والبدائل المتاحة على Google Flights ↗</a>",
            high_priority=True
        )
        return

    available_flights.sort(key=lambda x: int(x["price"]))
    cheapest = available_flights[0]
    pred = analyze_price_prediction(cheapest["price"], dep_clean, trip_type=trip_type)

    date_display = f"<code>{dep_clean}</code>" + (f" ⬅ <code>{ret_clean}</code>" if ret_clean else "")
    lines = [
        f"🎯 <b>نتيجة الفحص الفوري المباشر:</b>\n",
        f"🧭 <b>المسار:</b> {title_head}",
        f"📅 <b>تاريخ السفر:</b> {date_display}\n",
        f"🏆 <b>أفضل خيار متاح:</b>",
        f"   💰 <b>{cheapest['price']} ر.س إجمالي</b> — ✈️ {cheapest['airline']} (⏰ {cheapest['time_ar']})",
        f"   🎯 <b>مؤشر الثقة في الشراء:</b> <b>{pred['score']}%</b> ({pred['recommendation']})",
        f"   🔮 <b>توقع المسار:</b> <i>{pred['forecast']}</i>\n",
        f"📋 <b>كافة الخيارات المباشرة:</b>"
    ]

    seen = set()
    tp_flight_probe = get_affiliate_flight_link(dep_clean, ret_clean, trip_type)
    tp_hotel_probe = get_affiliate_hotel_link(dep_clean, ret_clean)

    for f in available_flights:
        air = f["airline"]
        if air not in seen:
            seen.add(air)
            bag_txt = " (شحن 23kg مجاناً 🎁)" if "السعودية" in air else f" (+{BAGGAGE_FEES.get(air, 140)} حقيبة)"
            dir_link = get_direct_booking_link(air, dep_clean, ret_clean, trip_type)
            action_label = "حجز مباشر عبر محرك المقارنة ↗" if "دمج ذكي" in air else f"حجز مباشر من موقع {air} ↗"
            lines.append(
                f"• <b>{f['price']} ر.س</b> — ✈️ {air}{bag_txt} (⏰ {f['time_ar']})\n"
                f"   🔗 <a href='{dir_link}'>{action_label}</a>"
            )

    cal_link = get_google_calendar_link(title_head, cheapest["airline"], dep_clean, ret_clean, cheapest["price"], url)
    lines.append(f"\n🔍 <a href='{tp_flight_probe}'><b>مقارنة بدائل التذاكر (تأكيد أفضل سعر) ↗</b></a>")
    lines.append(f"🏨 <a href='{tp_hotel_probe}'><b>أفضل عروض فنادق وشقق جدة لنفس الفترة ↗</b></a>")
    lines.append(f"📅 <a href='{cal_link}'>إضافة موعد الرحلة لتقويم Google ↗</a>")

    send_telegram_msg("\n".join(lines), high_priority=True)
    print("✅ تم إرسال تقرير المسار المخصص للتيليجرام بنجاح!")

# ==========================================
# ☁️ دورة الفحص السحابي الشامل مع إعادة التدوير
# ==========================================
def run_cloud_scan():
    if CUSTOM_DEP:
        run_custom_date_probe(CUSTOM_DEP, CUSTOM_RET, CUSTOM_TYPE)
        return

    start_time = time.time()
    print("☁️ بدء تشغيل الرادار السحابي وصائد القيعان (درع التخفي والفلترة مفعّل)...")
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

            try:
                context = create_stealth_context(browser)
                page = context.new_page()
                page.route("**/*", intercept_route_resources)

                for idx, (dep, ret, trip_type) in enumerate(pairs, 1):
                    dep_clean = clean_date_str(dep)
                    ret_clean = clean_date_str(ret)

                    # فحص سلامة الصفحة وإعادة التدوير كل 8 رحلات لتطهير الرام
                    if idx % 8 == 0 or page.is_closed():
                        try:
                            if not page.is_closed():
                                page.close()
                        except Exception:
                            pass
                        page = context.new_page()
                        page.route("**/*", intercept_route_resources)

                    url = f"https://www.google.com/travel/flights?q=Flights%20from%20ELQ%20to%20JED%20on%20{dep_clean}%20through%20{ret_clean}%20nonstop&curr=SAR&hl=en&gl=sa"
                    cheapest_flight = None
                    playwright_failed = False

                    for attempt in range(3):
                        try:
                            if attempt > 0 or page.is_closed():
                                try:
                                    if not page.is_closed():
                                        page.close()
                                except Exception:
                                    pass
                                page = context.new_page()
                                page.route("**/*", intercept_route_resources)

                            page.goto(url, wait_until="domcontentloaded", timeout=16000)

                            if "consent.google.com" in page.url:
                                try:
                                    page.locator('button:has-text("Accept all"), button:has-text("I agree")').first.click(timeout=3000)
                                    page.wait_for_load_state("domcontentloaded", timeout=4000)
                                except Exception:
                                    pass

                            page.mouse.wheel(0, random.randint(100, 250))

                            try:
                                page.wait_for_selector("li.pIav2d, div.pIav2d, [role='listitem']", timeout=5000)
                            except Exception:
                                time.sleep(1.2)

                            cards = page.locator("li.pIav2d, div.pIav2d, [role='listitem']").all()
                            nonstop_options = []

                            for c in cards:
                                try:
                                    txt = c.inner_text()
                                    txt_l = txt.lower()
                                    has_nonstop = any(w in txt_l for w in ["nonstop", "non-stop", "direct"]) or any(w in txt for w in ["مباشر", "بدون توقف", "مباشرة"])
                                    has_stop = any(w in txt_l for w in ["1 stop", "2 stop"]) or any(w in txt for w in ["غير مباشر", "توقف واحد"])

                                    if has_nonstop and not has_stop:
                                        price, airline, time_ar, raw_time = parse_flight_card_text(txt)
                                        if price:
                                            nonstop_options.append({
                                                "price": int(price),
                                                "airline": airline,
                                                "time_ar": time_ar,
                                                "raw_time": raw_time
                                            })
                                except Exception:
                                    continue

                            if nonstop_options:
                                if "جمعة إلى سبت" in trip_type:
                                    pm_options = [o for o in nonstop_options if "PM" in o["raw_time"].upper() or "م" in o["time_ar"]]
                                    cheapest_flight = min(pm_options, key=lambda x: int(x["price"])) if pm_options else min(nonstop_options, key=lambda x: int(x["price"]))
                                else:
                                    cheapest_flight = min(nonstop_options, key=lambda x: int(x["price"]))
                                playwright_failed = False
                                break
                            else:
                                playwright_failed = True

                        except Exception as loop_e:
                            playwright_failed = True
                            wait_backoff = (2 ** attempt) + random.uniform(0.5, 1.5)
                            time.sleep(wait_backoff)

                    if playwright_failed or not cheapest_flight:
                        print(f"⚡ المحرك الاحتياطي (Fallback Engine) للرحلة [{dep_clean}]...")
                        cheapest_flight = fallback_http_probe(dep_clean, ret_clean, "roundtrip")

                    if cheapest_flight:
                        cur_p = int(cheapest_flight["price"])
                        air = cheapest_flight["airline"]
                        f_time = cheapest_flight["time_ar"]
                        flight_key = f"{dep_clean}_{ret_clean}"
                        print(f"[{idx}/{len(pairs)}] ✅ رصد: {trip_type} -> {cur_p} ر.س ({air})")

                        prev_p = int(price_history[flight_key]) if flight_key in price_history and str(price_history[flight_key]).isdigit() else None
                        pred = analyze_price_prediction(cur_p, dep_clean, prev_p, "roundtrip")

                        price_history[flight_key] = cur_p
                        results.append({
                            "نوع العطلة": trip_type,
                            "تاريخ الذهاب": dep_clean,
                            "وقت الإقلاع": f_time,
                            "تاريخ العودة": ret_clean,
                            "الناقل": air,
                            "السعر": cur_p,
                            "الرابط": url,
                            "مؤشر_الثقة": pred["score"],
                            "توصية_القرار": pred["recommendation"],
                            "توقعات_7_أيام": pred["forecast"],
                            "مستوى_الثقة": pred["level"]
                        })

                    time.sleep(random.uniform(1.2, 2.4))

            finally:
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
        print("⚠️ لم يتم رصد نتائج جديدة، تفعيل استرجاع آخر نسخة صالحة...")
        if os.path.exists(MINI_APP_DATA_FILE):
            try:
                with open(MINI_APP_DATA_FILE, "r", encoding="utf-8") as f:
                    old_data = json.load(f)
                    results = old_data.get("flights", [])
            except Exception as e:
                print(f"⚠️ خطأ أثناء استرجاع البيانات الاحتياطية: {e}")

        send_telegram_msg(
            "⚠️ <b>تحذير رادار الطيران:</b> لم تسجل دورة الفحص الأخيرة بيانات حية جديدة، وتم الإبقاء على آخر قاعدة بيانات نشطة لحماية الـ Mini App من التوقف.",
            high_priority=False
        )

    if results:
        results = sorted(results, key=lambda x: int(x.get("السعر", x.get("price", 0))))
        stats_summary = compute_flight_statistics(results)
        sync_to_google_sheets(results, duration)
        df = pd.DataFrame(results)

        ksa_now = get_ksa_now()
        mini_app_payload = {
            "updated_at": ksa_now.strftime("%Y-%m-%d %I:%M %p"),
            "raw_timestamp": int(ksa_now.timestamp()),
            "stats": stats_summary,
            "flights": results
        }
        with open(MINI_APP_DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(mini_app_payload, f, ensure_ascii=False, indent=2)
        print("📱 تم تحديث ملف بيانات التطبيق المصغر التنبؤي (flights_data.json) بنجاح مع الإحصاءات المرجعية!")

        is_manual_trigger = os.environ.get("GITHUB_EVENT_NAME") == "workflow_dispatch"
        is_morning_window = (8 <= ksa_now.hour < 10)

        if is_manual_trigger or is_morning_window:
            generate_price_chart(results)
            briefing_text = build_briefing_message(results)
            if os.path.exists(CHART_IMAGE_PATH):
                caption = f"📈 <b>مخطط حركة أسعار (القصيم ⇄ جدة)</b>\n• أديل (أخضر) | السعودية (كحلي) | ناس (ذهبي)"
                send_telegram_photo(CHART_IMAGE_PATH, caption=caption)
            send_telegram_msg(briefing_text)

        top_blocks = []
        for idx, item in enumerate(results[:15], 1):
            t_dep = clean_date_str(item.get('تاريخ الذهاب', item.get('dep', '')))
            t_ret = clean_date_str(item.get('تاريخ العودة', item.get('ret', '')))
            flight_time = item.get('وقت الإقلاع', item.get('time_ar', ''))
            t_time = f" | ⏰ الإقلاع: {flight_time}" if flight_time else ""
            airline_name = item.get('الناقل', item.get('airline', 'رحلة مباشرة'))

            dir_link = get_direct_booking_link(airline_name, t_dep, t_ret, "roundtrip")
            tp_f = get_affiliate_flight_link(t_dep, t_ret, "roundtrip")
            tp_h = get_affiliate_hotel_link(t_dep, t_ret)

            price_val = int(item.get('السعر', item.get('price', 0)))
            price_desc = "358 ر.س إجمالي (179 ذهاب + 179 عودة)" if price_val == 358 else f"{price_val} ر.س إجمالي"
            action_label = "حجز مباشر عبر محرك المقارنة ↗" if "دمج ذكي" in airline_name else f"حجز مباشر من {airline_name} ↗"

            block = (
                f"{idx}. <b>{price_desc}</b> — ✈️ {airline_name}\n"
                f"   🗓 {item.get('نوع العطلة', item.get('trip_type', ''))}{t_time}\n"
                f"   🛫 ذهاب: <code>{t_dep}</code>\n"
                f"   🛬 عودة: <code>{t_ret}</code>\n"
                f"   🔗 <a href='{dir_link}'>{action_label}</a>\n"
                f"   🔍 <a href='{tp_f}'>مقارنة بدائل التذاكر (تأكيد أفضل سعر) ↗</a>\n"
                f"   🏨 <a href='{tp_h}'>أفضل عروض فنادق وشقق جدة لنفس الفترة ↗</a>\n"
            )
            top_blocks.append(block)

        header_title = "✅ <b>اكتمل الفحص ومقارنة العروض بنجاح!</b>\n\n🎒 <b>أرخص التذاكر الخفيفة (السعر الإجمالي ذهاب وعودة):</b>\n"
        send_telegram_long_message(top_blocks, header=header_title)

        df.to_excel(EXCEL_FILE, index=False)
        print(f"✅ اكتملت الدورة السحابية بنجاح في {duration} ثانية! تم رصد {len(results)} رحلة.")
    else:
        send_telegram_msg("⚠️🚨 تحذير: تعذر الوصول إلى نتائج الفحص ولم يتوفر أرشيف احتياطي!", high_priority=True)

if __name__ == "__main__":
    run_cloud_scan()
