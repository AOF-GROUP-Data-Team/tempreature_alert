# ============================================================
# TEMPERATURE CHECK REPORT — Template 943193
# Runs every 2 hours — 6 times daily
# GitHub Actions Version
# ============================================================

# !pip install requests pandas openpyxl Pillow gspread gspread-dataframe google-auth-oauthlib pytz google-api-python-client

from google.auth import default
import gspread
from gspread_dataframe import set_with_dataframe
import os
import requests
import pandas as pd
import io
import pytz
import json
from datetime import datetime, timedelta
from urllib.parse import quote
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.utils.units import pixels_to_EMU

# ── Auth ──────────────────────────────────────────────────
import json, os
from google.oauth2.service_account import Credentials as SACredentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
sa_info      = json.loads(os.environ.get('GOOGLE_CREDENTIALS'))
google_creds = SACredentials.from_service_account_info(sa_info, scopes=SCOPES)
gc           = gspread.authorize(google_creds)
print(f"✅ Authenticated as: {sa_info.get('client_email')}")

# ── CONFIG ────────────────────────────────────────────────
API_KEY         = os.environ.get("ZENPUT_API_KEY")
TEMPLATE_ID     = 943193
GOOGLE_SHEET_ID = "19hPHSr6dmFMshW2S4G6T16ZZkcL3B0SdkvkSDpQ11W4"
TZ              = pytz.timezone("Asia/Baghdad")

# ── TIME WINDOW ─────────────────────────────────────────
# Each run covers a 90-minute window ending at the run time.
# Schedule:  1:30am, 4:30am, 7:30am, 10:30am, 1:30pm, 4:30pm, 10:30pm KSA
# Window end   = now (rounded to nearest scheduled time)
# Window start = end - 90 minutes

_input_start = os.environ.get('INPUT_START', '').strip()
_input_end   = os.environ.get('INPUT_END',   '').strip()

if _input_start and _input_end:
    WINDOW_START = TZ.localize(datetime.strptime(_input_start, "%Y-%m-%d %H:%M"))
    WINDOW_END   = TZ.localize(datetime.strptime(_input_end,   "%Y-%m-%d %H:%M"))
    print(f"📅 Manual window: {WINDOW_START.strftime('%Y-%m-%d %H:%M')} → {WINDOW_END.strftime('%Y-%m-%d %H:%M')}")
else:
    WINDOW_END   = datetime.now(TZ)
    WINDOW_START = WINDOW_END - timedelta(minutes=90)
    print(f"📅 Auto window: {WINDOW_START.strftime('%Y-%m-%d %H:%M')} → {WINDOW_END.strftime('%Y-%m-%d %H:%M')}")

RUN_DATE = WINDOW_END.date()
DATE_STR = WINDOW_START.date().strftime("%Y-%m-%d")

# Photo settings
IMG_SIZE    = 250
IMG_HEIGHT  = 200
GAP_PX      = 6
PX_PER_CHAR = 7.0

# ── FIELD IDs ─────────────────────────────────────────────
FID_LOCATION       = 13660173
FID_PREPARED_BY    = 13660190
FID_DATETIME       = 13660193

# Chillers (photo + temp pairs)
FID_MAIN_CHILLER_PHOTO   = 13660194; FID_MAIN_CHILLER_TEMP   = 13660195
FID_MAIN_FREEZER_PHOTO   = 13660196; FID_MAIN_FREEZER_TEMP   = 13660197
FID_SHAWARMA_CHILLER_PHOTO = 13660198; FID_SHAWARMA_CHILLER_TEMP = 13660199
FID_ARABI_CHILLER_PHOTO  = 13660200; FID_ARABI_CHILLER_TEMP  = 13660201
FID_BURGER_CHILLER_PHOTO = 13660202; FID_BURGER_CHILLER_TEMP = 13660203
FID_MINI_FREEZER_PHOTO   = 13660204; FID_MINI_FREEZER_TEMP   = 13660205
FID_DRINKS_CHILLER_PHOTO = 13660397; FID_DRINKS_CHILLER_TEMP = 13660398

# New main freezer section
FID_MAIN_FREEZER2_PHOTO  = None   
FID_MAIN_FREEZER2_TEMP   = None

# Linear freezer (optional)
FID_LINEAR_FREEZER_PHOTO = 15390788; FID_LINEAR_TEMP = 15390790

# Temperature ranges for each chiller
TEMP_RANGES = {
    "Main Chiller":    (0,   5),
    "Main Freezer":    (-18, -16),
    "Shawarma Chiller":(0,   5),
    "Arabi Chiller":   (0,   5),
    "Burger Chiller":  (0,   5),
    "Drinks Chiller":  (0,   5),
    "Mini Freezer":    (-18, -16),
    "Linear Freezer":  (-18, -16),
}

# All chiller definitions in order
CHILLERS = [
    ("Main Chiller",     FID_MAIN_CHILLER_PHOTO,    FID_MAIN_CHILLER_TEMP),
    ("Main Freezer",     FID_MAIN_FREEZER_PHOTO,    FID_MAIN_FREEZER_TEMP),
    ("Shawarma Chiller", FID_SHAWARMA_CHILLER_PHOTO, FID_SHAWARMA_CHILLER_TEMP),
    ("Arabi Chiller",    FID_ARABI_CHILLER_PHOTO,   FID_ARABI_CHILLER_TEMP),
    ("Burger Chiller",   FID_BURGER_CHILLER_PHOTO,  FID_BURGER_CHILLER_TEMP),
    ("Drinks Chiller",   FID_DRINKS_CHILLER_PHOTO,  FID_DRINKS_CHILLER_TEMP),
    ("Mini Freezer",     FID_MINI_FREEZER_PHOTO,    FID_MINI_FREEZER_TEMP),
    ("Linear Freezer",   FID_LINEAR_FREEZER_PHOTO,  FID_LINEAR_TEMP),
]

BRANCH_MAP = {
    "2197299": "LBRUH B07",   "2239240": "FAYJED B32",  "2235670": "ANRUH B31",
    "2190657": "SLAHS B23",   "2164026": "NDRUH B15",   "2164019": "SWRUH B08",
    "2203271": "SARUH B27",   "2164017": "DARUH B06",   "2164032": "KRRUH B21",
    "2164031": "SFJED B24",   "2164025": "RBRUH B14",   "2164016": "RAWRUH B05",
    "2197297": "NSRUH B04",   "2164021": "SHRUH B10",   "2164013": "KHRUH B02",
    "2155652": "NURUH B01",   "2164023": "TWRUH B12",   "2164020": "AZRUH B09",
    "2199002": "RWAHS B25",   "2242934": "HIRJED B33",  "2164022": "NRRUH B11",
    "2164030": "MURUH B19",   "2164014": "GHRUH B03",   "2211854": "QADRUH B30",
    "2254072": "Garatiss QB03","2256386": "Garatiss QB04","2258220": "PSJED B36",
    "2169459": "Lubda Alaqeq Branch LB01",               "2232755": "Garatis As Suwaidi QB01",
    "2185452": "OBJED B22",   "2243963": "URURUH B34",  "2222802": "Lubda Alkhaleej Branch LB02",
    "2199835": "HAJED B26",   "2210205": "MAJED B28",   "2250799": "IRRUH B35",
    "2164027": "BDRUH B16",   "2155654": "AQRUH B13",   "2197298": "TKRUH B18",
    "2171883": "Twesste TW01","2235805": "Garatis Alnargis QB02",
    "2164028": "QRRUH B17",   "2230615": "Central Kitchen",
    "2256173": "Al Hair WH",  "2257790": "SHWMAK B37",  "2260889": "UHDMM B38",
    "2263062": "HSRUH B39",   "2268360": "QB05",
}

# ── HELPERS ───────────────────────────────────────────────
def zenput_headers():
    return {"X-API-TOKEN": API_KEY, "Accept": "application/json"}

def gv(ans_by_id, fid):
    if not fid:
        return ""
    ans = ans_by_id.get(fid)
    if not ans:
        return ""
    val = ans.get("value")
    if val is None:
        return ""
    if isinstance(val, list):
        if val and isinstance(val[0], dict) and "s3_key" in val[0]:
            return val   # return raw list for photo fields
        return ", ".join(str(v).strip() for v in val if v)
    return str(val).strip()

def get_s3_keys(ans_by_id, fid):
    val = gv(ans_by_id, fid)
    if isinstance(val, list):
        return [item["s3_key"] for item in val if isinstance(item, dict) and "s3_key" in item]
    return []

def get_signed_url(s3_key):
    url = f"https://www.zenput.com/api/v2/users/current/storage/?path={quote(s3_key)}"
    try:
        resp = requests.get(url, headers=zenput_headers(), timeout=10)
        if resp.status_code == 200:
            return resp.json().get("data", {}).get("location", "")
    except Exception as e:
        print(f"⚠️  Sign URL failed: {e}")
    return ""

# Freezer chillers that must always be negative
FREEZER_CHILLERS = {"Main Freezer", "Mini Freezer", "Linear Freezer"}

def normalize_temp(temp_str, chiller_name):
    """Parse temp value, auto-apply negative sign for freezers if forgotten."""
    try:
        val = float(str(temp_str).strip().replace("−", "-").replace("–", "-"))
        # Auto-fix: if freezer and positive, make it negative
        if chiller_name in FREEZER_CHILLERS and val > 0:
            print(f"  🔧 Auto-corrected: {chiller_name} {val} → {-val} (freezer must be negative)")
            val = -val
        return val
    except Exception:
        return None

def is_in_range(temp_str, chiller_name):
    val = normalize_temp(temp_str, chiller_name)
    if val is None:
        return True
    lo, hi = TEMP_RANGES.get(chiller_name, (None, None))
    if lo is None:
        return True
    return lo <= val <= hi

def temp_status(temp_str, chiller_name):
    """Returns (display_str, normalized_value, is_ok)"""
    if not temp_str or str(temp_str).strip() == "":
        return "—", None, True
    val = normalize_temp(temp_str, chiller_name)
    if val is None:
        return str(temp_str), None, True
    ok = is_in_range(temp_str, chiller_name)
    lo, hi = TEMP_RANGES.get(chiller_name, ("?", "?"))
    suffix = " ✅" if ok else f" 🚨 (Range: {lo}~{hi})"
    return f"{val}°C{suffix}", val, ok

def download_image(signed_url):
    try:
        resp = requests.get(signed_url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code != 200:
            return None
        img = PILImage.open(io.BytesIO(resp.content)).convert("RGB")
        img.thumbnail((IMG_SIZE, IMG_SIZE), PILImage.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        buf.seek(0)
        return XLImage(buf)
    except Exception as e:
        print(f"  ⚠️  Photo download failed: {e}")
        return None

# ── FETCH ─────────────────────────────────────────────────
def fetch_submissions():
    all_subs, start, limit = [], 0, 100
    fetch_from = WINDOW_START.strftime("%Y-%m-%d")
    print(f"🚀 Fetching template {TEMPLATE_ID} | window: {WINDOW_START.strftime('%H:%M')} → {WINDOW_END.strftime('%H:%M')}...")
    while True:
        resp = requests.get(
            "https://www.zenput.com/api/v3/submissions/",
            headers=zenput_headers(),
            params={"form_template_id": TEMPLATE_ID, "limit": limit,
                    "start": start, "date_submitted_start": fetch_from}
        )
        if resp.status_code != 200:
            print(f"⚠️  HTTP {resp.status_code} — stopping.")
            break
        batch = resp.json().get("data", [])
        if not batch:
            break
        stop_early = False
        for s in batch:
            date_raw = s.get("smetadata", {}).get("date_submitted_local", "")
            try:
                sub_dt = datetime.fromisoformat(date_raw)
                if sub_dt.tzinfo is None:
                    sub_dt = TZ.localize(sub_dt)
                if sub_dt < WINDOW_START:
                    stop_early = True
                    break
                if WINDOW_START <= sub_dt <= WINDOW_END:
                    all_subs.append(s)
            except Exception:
                pass
        if stop_early or len(batch) < limit:
            break
        start += limit
    print(f"✅ {len(all_subs)} submissions in window.")
    return all_subs

# ── PARSE ─────────────────────────────────────────────────
def parse_submissions(all_subs):
    rows = []
    print("⚙️  Parsing and signing photo URLs...")
    for s in all_subs:
        meta     = s.get("smetadata", {}) or {}
        answers  = s.get("answers", []) or []
        location = meta.get("location", {}) or {}
        loc_id   = str(location.get("id", ""))
        branch   = " ".join(BRANCH_MAP.get(loc_id, location.get("name", "Unknown").strip()).split())
        cb       = meta.get("created_by", {})
        submitter= cb.get("display_name", "") if isinstance(cb, dict) else str(cb or "")

        ans_by_id = {a.get("field_id"): a for a in answers if a.get("field_id")}

        # Use actual submission time from metadata
        dt_raw = meta.get("date_submitted_local", "")
        try:
            dt_obj = datetime.fromisoformat(str(dt_raw))
            dt_str = dt_obj.strftime("%Y-%m-%d %H:%M")
        except Exception:
            dt_str = str(dt_raw)[:16]

        row = {
            "Date":        DATE_STR,
            "Date & Time": dt_str,
            "Branch":      branch,

        }

        # Extract each chiller temp + sign photo URLs
        all_ok = True
        for chiller_name, photo_fid, temp_fid in CHILLERS:
            temp_raw = gv(ans_by_id, temp_fid)
            s3_keys  = get_s3_keys(ans_by_id, photo_fid)
            temp_display, temp_val, ok = temp_status(temp_raw, chiller_name)
            if not ok:
                all_ok = False

            # Sign photo URLs
            signed_urls = []
            for key in s3_keys:
                url = get_signed_url(key)
                if url:
                    signed_urls.append(url)

            row[f"{chiller_name} Temp"]   = temp_display
            row[f"{chiller_name} Raw"]    = temp_raw
            row[f"{chiller_name} Val"]    = temp_val
            row[f"{chiller_name} OK"]     = ok
            row[f"{chiller_name} Photos"] = signed_urls

        row["All OK"] = all_ok
        rows.append(row)
        status = "✅" if all_ok else "🚨"
        print(f"  {status} {branch} — {dt_str}")

    print(f"✅ Parsed {len(rows)} records.")
    return rows

# ── EXCEL ─────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="2C3E50")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
ALT_FILL     = PatternFill("solid", fgColor="F2F2F2")
RED_FILL     = PatternFill("solid", fgColor="FFCCCC")
GREEN_FILL   = PatternFill("solid", fgColor="CCFFCC")
ORANGE_FILL  = PatternFill("solid", fgColor="FFE5CC")
CENTER       = Alignment(vertical="center", horizontal="center", wrap_text=True)
THIN         = Side(style="thin", color="CCCCCC")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def build_excel(rows, output_path):
    wb = Workbook()

    # One sheet per chiller + one summary sheet
    # ── Summary sheet ──────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"

    sum_cols = [
        ("Date & Time", 18), ("Branch", 22),
        ("Main Chiller", 16), ("Main Freezer", 16),
        ("Shawarma Chiller", 18), ("Arabi Chiller", 16),
        ("Burger Chiller", 16), ("Drinks Chiller", 16),
        ("Mini Freezer", 16), ("Linear Freezer", 16),
        ("Status", 12),
    ]
    for ci, (h, w) in enumerate(sum_cols, start=1):
        cell = ws_sum.cell(row=1, column=ci, value=h)
        cell.fill = HEADER_FILL; cell.font = HEADER_FONT
        cell.alignment = CENTER; cell.border = BORDER
        ws_sum.column_dimensions[get_column_letter(ci)].width = w
    ws_sum.row_dimensions[1].height = 30
    ws_sum.freeze_panes = "A2"

    for ri, row in enumerate(rows):
        er        = ri + 2
        all_ok    = row["All OK"]
        base_fill = ALT_FILL if ri % 2 == 1 else None

        values = [
            row["Date & Time"], row["Branch"],
            row.get("Main Chiller Raw", ""),
            row.get("Main Freezer Raw", ""),
            row.get("Shawarma Chiller Raw", ""),
            row.get("Arabi Chiller Raw", ""),
            row.get("Burger Chiller Raw", ""),
            row.get("Drinks Chiller Raw", ""),
            row.get("Mini Freezer Raw", ""),
            row.get("Linear Freezer Raw", ""),
            "✅ OK" if all_ok else "🚨 Alert",
        ]
        ok_flags = [
            True, True,
            row.get("Main Chiller OK", True),
            row.get("Main Freezer OK", True),
            row.get("Shawarma Chiller OK", True),
            row.get("Arabi Chiller OK", True),
            row.get("Burger Chiller OK", True),
            row.get("Drinks Chiller OK", True),
            row.get("Mini Freezer OK", True),
            row.get("Linear Freezer OK", True),
            all_ok,
        ]
        for ci, (val, ok) in enumerate(zip(values, ok_flags), start=1):
            cell = ws_sum.cell(row=er, column=ci, value=val)
            cell.border = BORDER; cell.alignment = CENTER
            cell.font   = Font(name="Arial", size=10, bold=(not ok))
            if not ok:
                cell.fill = RED_FILL
            elif base_fill:
                cell.fill = base_fill
        ws_sum.row_dimensions[er].height = 22

    # ── One sheet per chiller with photos ──────────────────
    CHILLER_COLORS = {
        "Main Chiller":     "1A73E8",
        "Main Freezer":     "34495E",
        "Shawarma Chiller": "E67E22",
        "Arabi Chiller":    "27AE60",
        "Burger Chiller":   "C0392B",
        "Drinks Chiller":   "8E44AD",
        "Mini Freezer":     "2980B9",
        "Linear Freezer":   "16A085",
    }

    # Calculate max photos per chiller dynamically
    max_photos_per_chiller = {}
    for chiller_name, _, _ in CHILLERS:
        max_p = max((len(row.get(f"{chiller_name} Photos", [])) for row in rows), default=1)
        max_photos_per_chiller[chiller_name] = max(max_p, 1)

    for chiller_name, _, _ in CHILLERS:
        # Only create sheet if any submission has data for this chiller
        has_data = any(row.get(f"{chiller_name} Raw") for row in rows)
        if not has_data:
            continue

        color = CHILLER_COLORS.get(chiller_name, "2C3E50")
        max_photos     = max_photos_per_chiller.get(chiller_name, 1)
        photo_col_px   = (IMG_SIZE + GAP_PX) * max_photos + 10
        photo_col_width = photo_col_px / PX_PER_CHAR
        ws = wb.create_sheet(title=chiller_name[:31])

        # Title
        ws.merge_cells("A1:D1")
        tc = ws["A1"]
        lo, hi = TEMP_RANGES.get(chiller_name, ("?", "?"))
        tc.value     = f"{chiller_name} — Expected: {lo}°C to {hi}°C — {DATE_STR}"
        tc.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
        tc.fill      = PatternFill("solid", fgColor=color)
        tc.alignment = CENTER
        ws.row_dimensions[1].height = 35

        sheet_cols = [
            ("Date & Time",  18),
            ("Branch",       22),
            ("Temperature",  16),
            ("Photo",        photo_col_width),
        ]
        for ci, (h, w) in enumerate(sheet_cols, start=1):
            cell = ws.cell(row=2, column=ci, value=h)
            cell.fill      = PatternFill("solid", fgColor=color)
            cell.font      = HEADER_FONT
            cell.alignment = CENTER; cell.border = BORDER
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[2].height = 28
        ws.freeze_panes = "A3"

        PHOTO_COL = 4  # col 1=Date&Time, 2=Branch, 3=Temperature, 4=Photo

        for ri, row in enumerate(rows):
            er        = ri + 3
            temp_raw  = row.get(f"{chiller_name} Raw", "")
            temp_disp = row.get(f"{chiller_name} Temp", "—")
            temp_val  = row.get(f"{chiller_name} Val", None)
            ok        = row.get(f"{chiller_name} OK", True)
            urls      = row.get(f"{chiller_name} Photos", [])
            has_photo = bool(urls)
            base_fill = ALT_FILL if ri % 2 == 1 else None

            ws.row_dimensions[er].height = IMG_HEIGHT if has_photo else 22

            # Write data cells
            data_vals = [row["Date & Time"], row["Branch"], temp_disp]
            for ci, val in enumerate(data_vals, start=1):
                cell = ws.cell(row=er, column=ci, value=val)
                cell.border = BORDER; cell.alignment = CENTER
                cell.font   = Font(name="Arial", size=10, bold=(ci == 3 and not ok))
                if ci == 3 and not ok:
                    cell.fill = RED_FILL
                elif ci == 3 and ok and temp_raw:
                    cell.fill = GREEN_FILL
                elif base_fill:
                    cell.fill = base_fill

            # Photo cell
            photo_cell = ws.cell(er, PHOTO_COL)
            photo_cell.border = BORDER
            if base_fill: photo_cell.fill = base_fill

            if not has_photo:
                photo_cell.value     = "No photo"
                photo_cell.font      = Font(italic=True, color="999999", name="Arial", size=9)
                photo_cell.alignment = CENTER
            else:
                left_offset = 4
                for ph_idx, url in enumerate(urls):
                    print(f"    📸 {chiller_name} | {row['Branch']} | photo {ph_idx+1}...")
                    img_obj = download_image(url)
                    if img_obj:
                        col_offset_emu = pixels_to_EMU(left_offset)
                        anchor = TwoCellAnchor(editAs="oneCell")
                        anchor._from = AnchorMarker(
                            col=PHOTO_COL-1, colOff=col_offset_emu,
                            row=er-1,        rowOff=pixels_to_EMU(4)
                        )
                        anchor.to = AnchorMarker(
                            col=PHOTO_COL-1, colOff=col_offset_emu + pixels_to_EMU(IMG_SIZE),
                            row=er,          rowOff=0
                        )
                        img_obj.anchor = anchor
                        ws.add_image(img_obj)
                        left_offset += IMG_SIZE + GAP_PX

    wb.save(output_path)
    print(f"✅ Excel saved → {output_path}")


# ── EMAIL ─────────────────────────────────────────────────
def send_email(rows, excel_path):
    import smtplib
    from email.message import EmailMessage

    APP_PASSWORD  = os.environ.get('GMAIL_APP_PASSWORD')
    SENDER_EMAIL  = "aof.group.auto@gmail.com"
    RECIPIENTS_TO = ["o.salahaddin@aofgroup.com"]
    RECIPIENTS_CC = ["m.hejazi@aofgroup.com", "a.alsalem@aofgroup.com"]

    if not rows:
        subject = f"⚠️ لا توجد بيانات درجات الحرارة | من {WINDOW_START.strftime('%H:%M')} إلى {WINDOW_END.strftime('%H:%M')} ({DATE_STR})"
        body = f"""<html><head><meta charset="UTF-8"></head>
        <body dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;padding:20px;">
            <div style="background:#2C3E50;color:#fff;padding:16px;border-radius:8px;margin-bottom:20px;">
                <h2 style="margin:0;">🌡️ تقرير درجات الحرارة</h2>
                <p style="margin:4px 0 0;opacity:0.8;">
                    <span dir="ltr">{WINDOW_START.strftime('%H:%M')} &rarr; {WINDOW_END.strftime('%H:%M')}</span> | <span dir="ltr">{DATE_STR}</span>
                </p>
            </div>
            <div style="background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:16px;">
                <p style="margin:0;">⚠️ لا توجد تقارير مسجلة لهذه الفترة.</p>
            </div>
        </body></html>"""
    else:
        # Build summary table
        alert_rows = [r for r in rows if not r["All OK"]]
        ok_rows    = [r for r in rows if r["All OK"]]

        def make_rows(row_list, bg):
            html = ""
            for r in row_list:
                alerts = []
                for cn, _, _ in CHILLERS:
                    if not r.get(f"{cn} OK", True):
                        val = r.get(f"{cn} Raw", "")
                        lo, hi = TEMP_RANGES.get(cn, ("?","?"))
                        alerts.append(f"{cn}: {val}°C (المطلوب: {lo}~{hi})")
                alert_str = "<br>".join(alerts) if alerts else "✅"
                icon = "🚨" if alerts else "✅"
                html += f"""<tr style="background:{bg};">
                    <td style="padding:8px 12px;border:1px solid #ddd;">{icon} {r['Branch']}</td>
                    <td style="padding:8px 12px;border:1px solid #ddd;">{r['Date & Time']}</td>
                    <td style="padding:8px 12px;border:1px solid #ddd;font-size:12px;">{alert_str}</td>
                </tr>"""
            return html

        subject = f"🌡️ تقرير درجات الحرارة {'🚨 تنبيهات' if alert_rows else '✅ سليم'} | من {WINDOW_START.strftime('%H:%M')} إلى {WINDOW_END.strftime('%H:%M')}"

        body = f"""<html><head><meta charset="UTF-8"></head>
        <body dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;line-height:1.8;color:#000;max-width:900px;margin:auto;padding:20px;">
            <div style="background:#2C3E50;color:#fff;padding:16px 20px;border-radius:8px;margin-bottom:20px;">
                <h2 style="margin:0;">🌡️ تقرير درجات الحرارة</h2>
                <p style="margin:4px 0 0;font-size:13px;opacity:0.8;">
                    الفترة: <span dir="ltr">{WINDOW_START.strftime('%H:%M')} &rarr; {WINDOW_END.strftime('%H:%M')}</span> |
                    {len(rows)} فرع | {len(alert_rows)} تنبيه
                </p>
            </div>

            {"" if not alert_rows else f'''
            <div style="background:#FFCCCC;border:1px solid #CC0000;border-radius:8px;padding:12px 16px;margin-bottom:16px;">
                <p style="margin:0;font-weight:bold;color:#CC0000;">🚨 {len(alert_rows)} فرع خارج النطاق المسموح!</p>
            </div>'''}

            <table style="border-collapse:collapse;width:100%;margin-bottom:20px;font-size:13px;">
                <thead style="background:#2C3E50;color:#fff;">
                    <tr>
                        <th style="padding:10px 12px;border:1px solid #ddd;text-align:right;">الفرع</th>
                        <th style="padding:10px 12px;border:1px solid #ddd;text-align:right;">وقت التسجيل</th>
                        <th style="padding:10px 12px;border:1px solid #ddd;text-align:right;">الملاحظات</th>
                    </tr>
                </thead>
                <tbody>
                    {make_rows(alert_rows, "#FFEEEE")}
                    {make_rows(ok_rows,    "#FFFFFF")}
                </tbody>
            </table>

            <p style="color:#555;font-size:12px;">الملف المرفق يحتوي على التفاصيل الكاملة مع الصور.</p>
            <p style="color:#888;font-size:12px;">Business Intelligence — AOF Group</p>
        </body></html>"""

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"]    = f"Business Intelligence <{SENDER_EMAIL}>"
    msg["To"]      = ", ".join(RECIPIENTS_TO)
    msg["Cc"]      = ", ".join(RECIPIENTS_CC)
    msg.add_alternative(body, subtype="html")

    if excel_path and os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               filename=os.path.basename(excel_path))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)
    print("✅ Email sent.")

# ── GOOGLE SHEET ──────────────────────────────────────────
def write_to_sheet(rows):
    if not rows:
        return
    sheet_rows = []
    for row in rows:
        r = {
            "Date":        row["Date"],
            "Date & Time": row["Date & Time"],
            "Branch":      row["Branch"],
        }
        for chiller_name, _, _ in CHILLERS:
            r[chiller_name] = row.get(f"{chiller_name} Raw", "")
        r["Status"] = "OK" if row["All OK"] else "ALERT"
        sheet_rows.append(r)

    df = pd.DataFrame(sheet_rows)
    spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = spreadsheet.worksheet("Temperature Log")
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="Temperature Log", rows=5000, cols=20)
        set_with_dataframe(ws, df, row=1, col=1, include_index=False, include_column_header=True)
        print("✅ Sheet created.")
        return

    existing = ws.get_all_values()
    if len(existing) <= 1:
        set_with_dataframe(ws, df, row=1, col=1, include_index=False, include_column_header=True)
    else:
        ws.append_rows(df.fillna("").values.tolist(), value_input_option="USER_ENTERED")
    print(f"✅ {len(rows)} rows written to Google Sheet.")

# ── MAIN ──────────────────────────────────────────────────
print(f"📅 Running for {DATE_STR}")

all_subs = fetch_submissions()

if not all_subs:
    print("⚠️  No submissions found for today. Sending empty alert email...")
    send_email([], None)
else:
    rows = parse_submissions(all_subs)

    print("\n📊 Writing to Google Sheet...")
    write_to_sheet(rows)

    print("\n📊 Building Excel...")
    OUTPUT = f"Temperature_Report_{WINDOW_START.strftime('%Y-%m-%d_%H%M')}_{WINDOW_END.strftime('%H%M')}.xlsx"
    build_excel(rows, OUTPUT)

    print("\n📧 Sending email...")
    send_email(rows, OUTPUT)

print("🎉 Done!")
