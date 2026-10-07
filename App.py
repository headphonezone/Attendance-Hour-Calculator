import io
import re
import difflib
from datetime import datetime, date, time as dtime
from collections import defaultdict
import calendar

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

C_HEADER_BG  = "1F4E79"
C_HEADER_FG  = "FFFFFF"
C_SUBHDR_BG  = "2E75B6"
C_EXCESS_BG  = "E2EFDA"
C_SHORT_BG   = "FCE4D6"
C_NEUTRAL_BG = "DEEAF1"
C_ALT_ROW    = "F2F7FB"
C_BORDER     = "BDD7EE"
C_HOLIDAY_BG = "FFF2CC"
C_WFH_BG     = "E2F0D9"
C_RAW_HDR    = "375623"

def make_border(color=C_BORDER):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def make_font(bold=False, color="000000", size=10):
    return Font(name='Arial', bold=bold, color=color, size=size)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_align(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def style(cell, *, bold=False, fg="000000", size=10,
          fill=C_ALT_ROW, align_h='center', wrap=False, border=True):
    cell.font      = make_font(bold, fg, size)
    cell.fill      = make_fill(fill)
    cell.alignment = make_align(align_h, wrap=wrap)
    if border:
        cell.border = make_border()

def parse_punches(cell_value):
    if not cell_value:
        return []
    return [t.strip() for t in str(cell_value).split('\n') if t.strip()]

def minutes_to_hhmm(total_minutes):
    h, m = int(total_minutes // 60), int(total_minutes % 60)
    return f"{h}.{m:02d}"

def decimal_to_hhmm(decimal_hours):
    return minutes_to_hhmm(round(decimal_hours * 60))

# Excel stores time as a fraction of a 24-hour day. To show real hour
# totals (which routinely exceed 24) as proper H:MM — 60 minutes to the
# hour, not 100 — we store hours/24 and format the cell as elapsed time.
# Excel cannot render a negative duration in this format at all (always
# shows ####), so any value that can go negative (e.g. Net Hours) must be
# built as a text label via TEXT()/"-" concatenation instead.
TIME_FMT = "[h]:mm"

def to_excel_time(decimal_hours):
    return round(decimal_hours, 4) / 24

def compute_hours_from_pair(t_in_str, t_out_str):
    try:
        t_in  = datetime.strptime(t_in_str,  "%H:%M")
        t_out = datetime.strptime(t_out_str, "%H:%M")
        diff  = int((t_out - t_in).total_seconds() / 60)
        if diff <= 0:
            return 0.0, "0.00"
        return round(diff / 60, 4), minutes_to_hhmm(diff)
    except ValueError:
        return 0.0, "0.00"

# Part-timers work one of two fixed half-day shifts: a morning shift
# (09:30-13:30) or an evening shift (14:00-18:00). Full-timers work a
# single 09:30-18:00 shift. When only one punch is recorded for a day,
# infer the missing In/Out using whichever shift the punch falls in.
PT_MORNING_SHIFT = ("09:30", "13:30")
PT_EVENING_SHIFT = ("14:00", "18:00")
FT_SHIFT         = ("09:30", "18:00")

def infer_missing_punch(punch_str, is_part_time):
    try:
        t = datetime.strptime(punch_str, "%H:%M")
    except ValueError:
        return punch_str, punch_str

    if is_part_time:
        if t.hour < 14:
            shift_start, shift_end = PT_MORNING_SHIFT
        else:
            shift_start, shift_end = PT_EVENING_SHIFT
        midpoint = datetime.strptime(shift_start, "%H:%M") + (
            datetime.strptime(shift_end, "%H:%M") - datetime.strptime(shift_start, "%H:%M")
        ) / 2
    else:
        shift_start, shift_end = FT_SHIFT
        midpoint = datetime.strptime("12:00", "%H:%M")

    if t <= midpoint:
        return punch_str, shift_end     # looks like an In punch — default the Out
    else:
        return shift_start, punch_str   # looks like an Out punch — default the In

def get_week_number(day, year, month):
    wc, fw = 1, date(year, month, 1).weekday()
    for d in range(1, day + 1):
        if date(year, month, d).weekday() == 0 and d != 1:
            if not (d == 2 and fw == 6):
                wc += 1
    return wc

def get_week_target(relative_wk, year, month, daily_target):
    wd = 0
    for d_int in range(1, 32):
        try:
            if get_week_number(d_int, year, month) == relative_wk:
                wd += 1
        except ValueError:
            break
    return round(wd * daily_target, 2)

def get_month_sundays(year, month):
    total = calendar.monthrange(year, month)[1]
    return [d for d in range(1, total + 1) if date(year, month, d).weekday() == 6]

# ── ID normalization ────────────────────────────────────────────────────
# Excel stores IDs as numbers unless the cell is formatted as text, so the
# same employee ID can round-trip as "91", "91.0", " 91 " or "ABC01" vs
# "abc01" depending on which sheet it came from. Salary matching is by ID,
# so every ID must funnel through this before being stored or compared.
def normalize_id(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        try:
            f = float(s)
            if f == int(f):
                s = str(int(f))
        except ValueError:
            pass
    # Purely numeric IDs: strip leading zeros so "007" (kept as text in one
    # sheet) matches "7" (typed as a plain number in the other) — Excel
    # can't preserve leading zeros in a numeric cell, so the padded form
    # only ever shows up on one side.
    if s.isdigit():
        s = str(int(s))
    return s.upper()

# ── Salary Master parsing ──────────────────────────────────────────────
# Reusable file (ID, Name, Salary columns, header row optional) so the
# user only maintains one small sheet and re-uploads it every month
# instead of retyping salaries for every employee each time.
def _parse_salary_number(val):
    """Coerce a salary cell to a float. Handles plain numbers as well as
    text values with commas/currency symbols (e.g. "15,000", "Rs. 15000")
    that a straight float() call would reject."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r"[^\d.\-]", "", s)
    if not s or s in ("-", "."):
        return None
    try:
        return float(s)
    except ValueError:
        return None

def parse_salary_file(uploaded_file):
    salary_map, name_map, skipped = {}, {}, []
    fname = (uploaded_file.name or "").lower()
    try:
        if fname.endswith(".csv"):
            import csv as _csv
            content = uploaded_file.getvalue().decode("utf-8-sig")
            reader  = _csv.reader(content.splitlines())
            rows    = list(reader)
        else:
            # data_only=True: if a Salary cell holds a formula (e.g. a rate
            # lookup), read its last-saved computed value instead of the
            # formula text, which would otherwise fail number parsing.
            wb   = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
            ws   = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))

        for row_num, row in enumerate(rows, 1):
            if not row or all(c is None for c in row):
                continue
            row = list(row) + [None] * max(0, 3 - len(row))
            emp_id_raw, emp_name, salary_raw = row[0], row[1], row[2]

            if emp_id_raw is None:
                skipped.append((row_num, "—", "missing ID"))
                continue
            emp_id = normalize_id(emp_id_raw)
            if not emp_id or emp_id.lower() in ("id", "employee id"):
                continue

            salary = _parse_salary_number(salary_raw)
            if salary is None:
                skipped.append((row_num, emp_id, f"unreadable salary value: {salary_raw!r}"))
                continue
            if salary <= 0:
                skipped.append((row_num, emp_id, "salary is 0 or negative"))
                continue

            salary_map[emp_id] = salary
            name_map[emp_id]   = str(emp_name).strip() if emp_name else ""
    except Exception as e:
        skipped.append(("?", "?", f"file read error: {e}"))
    return salary_map, name_map, skipped

def make_salary_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salary Master"
    ws.append(["ID", "Name", "Monthly Salary"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Part-Time Master parsing ────────────────────────────────────────────
# Reusable file (ID, Name columns) so part-time employees don't need to be
# re-selected from a dropdown every month — re-upload the same file.
def parse_id_name_file(uploaded_file):
    id_set, name_map, skipped = set(), {}, []
    fname = (uploaded_file.name or "").lower()
    try:
        if fname.endswith(".csv"):
            import csv as _csv
            content = uploaded_file.getvalue().decode("utf-8-sig")
            reader  = _csv.reader(content.splitlines())
            rows    = list(reader)
        else:
            wb   = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
            ws   = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))

        for row_num, row in enumerate(rows, 1):
            if not row or all(c is None for c in row):
                continue
            row = list(row) + [None] * max(0, 2 - len(row))
            emp_id_raw, emp_name = row[0], row[1]

            if emp_id_raw is None:
                skipped.append((row_num, "—", "missing ID"))
                continue
            emp_id = normalize_id(emp_id_raw)
            if not emp_id or emp_id.lower() in ("id", "employee id"):
                continue

            id_set.add(emp_id)
            name_map[emp_id] = str(emp_name).strip() if emp_name else ""
    except Exception as e:
        skipped.append(("?", "?", f"file read error: {e}"))
    return id_set, name_map, skipped

def make_parttime_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Part-Time Master"
    ws.append(["ID", "Name"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def parse_logs_sheet(ws):
    all_rows   = list(ws.iter_rows(values_only=True))
    period_str = ""
    year, month = datetime.now().year, datetime.now().month

    for row in all_rows[:5]:
        for cell in row:
            if cell and isinstance(cell, str) and '~' in cell:
                period_str = cell.strip()
                try:
                    dt          = datetime.strptime(period_str.split('~')[0].strip(), "%Y/%m/%d")
                    year, month = dt.year, dt.month
                except Exception:
                    pass

    raw_records, emp_order = {}, []
    i = 0
    while i < len(all_rows):
        row = all_rows[i]
        if row and row[0] == 'No :':
            emp_no    = normalize_id(row[2]) if row[2] else 'Unknown'
            emp_name  = str(row[10]).strip() if row[10] else 'Unnamed'
            uid       = f"{emp_name.title()} (ID: {emp_no})"
            days_row  = all_rows[i - 1] if i > 0 else []
            pr_idx    = i + 1
            if pr_idx < len(all_rows):
                punch_row   = all_rows[pr_idx]
                day_punches = {}
                for col, day_num in enumerate(days_row):
                    if not isinstance(day_num, int): continue
                    if col >= len(punch_row): continue
                    try:
                        date(year, month, day_num)
                    except:
                        continue
                    punches = parse_punches(punch_row[col])
                    if punches:
                        day_punches[day_num] = punches
                if uid not in raw_records:
                    raw_records[uid] = {'name': emp_name, 'id': emp_no, 'punches': {}}
                    emp_order.append(uid)
                raw_records[uid]['punches'].update(day_punches)
        i += 1
    return raw_records, emp_order, period_str, year, month

# ── Filled attendance workbook (replaces the raw Logs sheet) ────────────
# The filled workbook holds three sheets we care about, found by content
# rather than by name since the tab names ("Sheet1", "Sheet4"...) are not
# stable month to month:
#   • attendance sheet — one In/Out block per employee (names only, no IDs),
#     a daily-target row above each block, yellow = Sunday/holiday
#   • logs sheet       — the raw device export; used only for IDs
#   • salary sheet     — "Name of the Staff" / "Basic Salary" list
NAME_MATCH_THRESHOLD = 0.85   # below this the match is left for the user to pick

def _time_str(v):
    if isinstance(v, (datetime, dtime)):
        return v.strftime("%H:%M")
    if isinstance(v, (int, float)) and 0 <= v < 1:
        m = round(v * 1440)
        return f"{m // 60:02d}:{m % 60:02d}"
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%H:%M", "%H:%M:%S"):
            try:
                return datetime.strptime(s, fmt).strftime("%H:%M")
            except ValueError:
                pass
    return None

def _time_hours(v):
    s = _time_str(v)
    if not s:
        return None
    h, m = s.split(":")
    return int(h) + int(m) / 60

def _is_yellow(cell):
    f = cell.fill
    return (f is not None and f.fill_type == "solid"
            and f.fgColor.type == "rgb"
            and str(f.fgColor.rgb).upper().endswith("FFFF00"))

def _norm_name(s):
    s = re.sub(r"\(.*?\)", " ", str(s).lower())
    s = re.sub(r"\b(w\.?h|sweeper)\b", " ", s)
    return re.sub(r"[^a-z]", "", s)

def _name_score(a, b):
    a, b = _norm_name(a), _norm_name(b)
    if not a or not b:
        return 0.0
    r = difflib.SequenceMatcher(None, a, b).ratio()
    # Device names are truncated ("gayathrim", "venkataleksh"), so a clean
    # prefix match is as good as an exact one.
    if min(len(a), len(b)) >= 4 and (a.startswith(b) or b.startswith(a)):
        r = max(r, 0.9)
    return r

def match_names(sources, candidates, threshold=NAME_MATCH_THRESHOLD):
    """One-to-one fuzzy match. Returns, per source, the index of its
    candidate or None. Best scores are claimed first so a weak match can
    never steal a candidate from a strong one."""
    pairs = sorted(
        ((_name_score(s, c), i, j) for i, s in enumerate(sources)
                                    for j, c in enumerate(candidates)),
        reverse=True,
    )
    result, used = [None] * len(sources), set()
    for score, i, j in pairs:
        if score < threshold:
            break
        if result[i] is None and j not in used:
            result[i] = j
            used.add(j)
    return result

def parse_attendance_sheet(ws):
    """Returns (blocks, year, month, holiday_days). Each block is
    {'name', 'target', 'punches': {day: [in, out]}} where punches use the
    same "HH:MM" strings the old Logs parser produced."""
    rows = list(ws.iter_rows())
    blocks, holiday_days = [], set()
    year = month = None

    for idx, row in enumerate(rows):
        label = row[1].value if len(row) > 1 else None
        if not (isinstance(label, str) and label.strip().lower() == "in time"):
            continue
        if idx < 3 or idx + 1 >= len(rows):
            continue
        in_row, out_row = row, rows[idx + 1]
        date_row, day_row, target_row = rows[idx - 1], rows[idx - 2], rows[idx - 3]

        if year is None and isinstance(date_row[1].value, datetime):
            year, month = date_row[1].value.year, date_row[1].value.month

        name_cell = in_row[0].value or out_row[0].value
        name      = " ".join(str(name_cell).split()) if name_cell else ""

        punches, targets, yellow = {}, [], []
        for col in range(2, len(in_row)):
            day = day_row[col].value if col < len(day_row) else None
            if not isinstance(day, int):
                continue
            t = _time_hours(target_row[col].value) if col < len(target_row) else None
            if t:
                targets.append(t)
            if _is_yellow(in_row[col]):
                yellow.append(day)
            t_in  = _time_str(in_row[col].value)
            t_out = _time_str(out_row[col].value) if col < len(out_row) else None
            day_p = [t for t in (t_in, t_out) if t]
            if day_p:
                punches[day] = day_p

        if year is not None:
            for d in yellow:
                try:
                    if date(year, month, d).weekday() != 6:   # Sundays are auto-credited
                        holiday_days.add(d)
                except ValueError:
                    pass

        target = max(set(targets), key=targets.count) if targets else None
        blocks.append({'name': name, 'target': target, 'punches': punches})

    return blocks, year, month, holiday_days

def parse_salary_sheet(ws):
    """[(name, monthly_salary)] from a "Name of the Staff" / "Basic Salary"
    sheet. Reads cached values, so formula cells need the workbook to have
    been saved by Excel."""
    rows = list(ws.iter_rows(values_only=True))
    name_col = sal_col = hdr = None
    for i, row in enumerate(rows[:15]):
        low = [str(c).strip().lower() if c is not None else "" for c in row]
        if "name of the staff" in low:
            name_col = low.index("name of the staff")
            sal_col  = low.index("basic salary") if "basic salary" in low else None
            hdr = i
            break
    if hdr is None or sal_col is None:
        return []
    out = []
    for row in rows[hdr + 1:]:
        if len(row) <= max(name_col, sal_col) or not row[name_col]:
            continue
        sal = _parse_salary_number(row[sal_col])
        if sal and sal > 0:
            out.append((" ".join(str(row[name_col]).split()), sal))
    return out

def _find_sheets(wb):
    att = logs = sal = None
    for ws in wb.worksheets:
        head = list(ws.iter_rows(min_row=1, max_row=60, max_col=3, values_only=True))
        if att is None and any(isinstance(r[1], str) and r[1].strip().lower() == "in time"
                               for r in head if len(r) > 1):
            att = ws
        elif logs is None and any(r and r[0] == 'No :' for r in head):
            logs = ws
        elif sal is None and any("name of the staff" in
                                 [str(c).strip().lower() for c in r if c is not None]
                                 for r in ws.iter_rows(min_row=1, max_row=15, max_col=4, values_only=True)):
            sal = ws
    return att, logs, sal

@st.cache_data(show_spinner="Reading workbook…")
def load_filled_workbook(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    att, logs, sal = _find_sheets(wb)
    if att is None:
        return None
    blocks, year, month, holiday_days = parse_attendance_sheet(att)
    log_ids, period_str = [], ""
    if logs is not None:
        log_records, log_order, period_str, log_year, log_month = parse_logs_sheet(logs)
        log_ids = [(log_records[u]['id'], log_records[u]['name']) for u in log_order]
        if year is None:
            year, month = log_year, log_month
    salaries = parse_salary_sheet(sal) if sal is not None else []
    return {
        'blocks': blocks, 'year': year, 'month': month,
        'holiday_days': sorted(holiday_days), 'log_ids': log_ids,
        'period_str': period_str, 'salaries': salaries,
        'has_logs': logs is not None, 'has_salary': sal is not None,
    }

# ── CHANGE 1 + 2: Skip relieved employees + inject paid holiday hrs ──────
def build_employees_dec(emp_order, raw_records, fixes, wfh_records, year, month,
                         holiday_dates, part_time_list=None, pt_daily_target=4.0,
                         daily_target=8.5, ft_holiday_hours=8.5, pt_holiday_hours=8.0):
    employees_dec    = {}
    holiday_day_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    month_sundays    = set(get_month_sundays(year, month))
    part_time_list    = part_time_list or []

    for uid in emp_order:
        p_dict  = raw_records[uid]['punches']
        has_wfh = bool(wfh_records.get(uid, {}))
        # Paid Sunday/holiday credit uses its own fixed standard hours
        # (8:30 = 8.5 decimal FT / 8:00 = 8.0 decimal PT by default) — NOT
        # the attendance daily target (daily_target/pt_daily_target) and
        # NOT the payroll Sal Per Hour divisor (8.3 FT / 8.0 PT), which are
        # separate, independently configurable figures.
        emp_daily_target = pt_holiday_hours if uid in part_time_list else ft_holiday_hours

        # Skip relieved employees — no punches and no WFH for the entire month
        if not p_dict and not has_wfh:
            continue

        f_dict    = fixes.get(uid, {})
        week_data = defaultdict(dict)

        for day, p in p_dict.items():
            if day in f_dict:
                dec, _ = compute_hours_from_pair(f_dict[day]['in'], f_dict[day]['out'])
            elif len(p) >= 2:
                # Take the earliest punch as In and the latest as Out.
                # With exactly 2 punches this is just that pair; with 3+
                # (duplicate scans, stray mid-day punches, etc.) it avoids
                # misreading them as separate in/out sessions — e.g. a
                # duplicate first punch [13:25, 13:25, 18:15] used to pair
                # off as (13:25->13:25)=0hrs while dropping 18:15 entirely.
                dec, _ = compute_hours_from_pair(p[0], p[-1])
            elif len(p) == 1:
                t_in, t_out = infer_missing_punch(p[0], uid in part_time_list)
                dec, _ = compute_hours_from_pair(t_in, t_out)
            else:
                dec = 0.0
            if dec > 0:
                week_data[get_week_number(day, year, month)][day] = dec

        # Inject WFH hours for days with no punch data
        for day, info in wfh_records.get(uid, {}).items():
            hrs = info.get('hours', 0.0)
            if hrs > 0:
                wk = get_week_number(day, year, month)
                if day not in week_data.get(wk, {}):
                    week_data[wk][day] = hrs

        # Inject paid holiday hrs (at the employee's own daily target) —
        # credited to hours but NOT a working day. Applies to Sundays too.
        # Only fills days with no punch/WFH entry.
        for hday in holiday_day_nums:
            try:
                date(year, month, hday)
            except:
                continue
            wk = get_week_number(hday, year, month)
            if hday not in week_data.get(wk, {}):
                week_data[wk][hday] = emp_daily_target   # paid holiday

        # Auto-credit Sundays with no actual punch/WFH data — treated like a
        # paid day off (same as a holiday) so they never register as shortage.
        # Real Sunday punches (already added above) are left untouched.
        for sday in month_sundays:
            if sday in holiday_day_nums:
                continue   # already handled by holiday injection
            wk = get_week_number(sday, year, month)
            if sday not in week_data.get(wk, {}):
                week_data[wk][sday] = emp_daily_target

        if week_data:
            employees_dec[uid] = dict(week_data)
    return employees_dec

def get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records):
    total_days   = calendar.monthrange(year, month)[1]
    punched_days = set(raw_records[uid]['punches'].keys())
    holiday_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    wfh_days     = set(wfh_records.get(uid, {}).keys())
    sundays      = set(get_month_sundays(year, month))
    leave = 0
    for d in range(1, total_days + 1):
        # Sundays are auto-credited like a paid day off, so an unpunched
        # Sunday is not a leave day.
        if d in holiday_nums or d in wfh_days or d in sundays: continue
        if d not in punched_days:
            leave += 1
    return leave

def get_leave_days_by_week(uid, raw_records, year, month, holiday_dates, wfh_records):
    """Same definition of 'leave day' as get_leave_days, grouped by week
    number, so a week's target can be reduced by the leave taken in it —
    a leave day should count only as leave, not also as shortage."""
    total_days   = calendar.monthrange(year, month)[1]
    punched_days = set(raw_records[uid]['punches'].keys())
    holiday_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    wfh_days     = set(wfh_records.get(uid, {}).keys())
    sundays      = set(get_month_sundays(year, month))
    by_week = defaultdict(int)
    for d in range(1, total_days + 1):
        if d in holiday_nums or d in wfh_days or d in sundays: continue
        if d not in punched_days:
            by_week[get_week_number(d, year, month)] += 1
    return dict(by_week)

def get_effective_week_target(wk, year, month, daily_target, leave_by_week):
    """Weekly target minus the daily target for each leave day taken that
    week, so leave is reflected once (as leave) instead of twice (leave
    and shortage for the same day)."""
    base       = get_week_target(wk, year, month, daily_target)
    leave_days = leave_by_week.get(wk, 0)
    return round(max(0.0, base - leave_days * daily_target), 2)

def get_holidays_on_leave(uid, raw_records, year, month, holiday_dates, wfh_records):
    punched_days = set(raw_records[uid]['punches'].keys())
    wfh_days     = set(wfh_records.get(uid, {}).keys())
    count = 0
    for hd in holiday_dates:
        if hd.year != year or hd.month != month:
            continue
        d = hd.day
        if d in wfh_days:   continue
        if d not in punched_days:
            count += 1
    return count

# ── CHANGE 2b: get_days_worked — holidays & auto-credited Sundays NOT counted
# as working days. Based on actual raw punch/WFH data, not injected hours,
# so an auto-credited (unpunched) Sunday never counts as a day worked —
# only a Sunday with a real punch does.
def get_days_worked(uid, raw_records, wfh_records, holiday_dates, year, month):
    holiday_day_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    punched_days     = {d for d, p in raw_records[uid]['punches'].items() if p} - holiday_day_nums
    wfh_non_holiday  = {d for d in wfh_records.get(uid, {}) if d not in holiday_day_nums}
    return len(punched_days | wfh_non_holiday)

def sum_week_hours(day_dict):
    return sum(day_dict.values())

RAW_SHEET          = "_RawData"
RAW_DATA_START_ROW = 2

def write_raw_data_sheet(wb, employees_dec, emp_order, raw_records,
                          year, month, daily_target, part_time_list, pt_daily_target,
                          period_str, holiday_dates, wfh_records):
    ws = wb.create_sheet(RAW_SHEET)
    ws.sheet_state = 'hidden'

    for col, h in enumerate(["UID_KEY", "ID", "Name", "Week", "HoursWorked",
                              "Target", "Excess", "Shortage", "RawTarget"], 1):
        ws.cell(row=1, column=col, value=h)

    row     = RAW_DATA_START_ROW
    row_map = {}

    for uid in emp_order:
        if uid not in employees_dec:
            continue
        current_daily = pt_daily_target if uid in part_time_list else daily_target
        week_dict     = employees_dec[uid]
        row_map[uid]  = {}
        leave_by_week = get_leave_days_by_week(uid, raw_records, year, month,
                                                holiday_dates, wfh_records)

        for wk in sorted(week_dict.keys()):
            wk_target     = get_effective_week_target(wk, year, month, current_daily, leave_by_week)
            wk_target_raw = get_week_target(wk, year, month, current_daily)
            wk_hrs_dec    = sum_week_hours(week_dict[wk])

            ws.cell(row=row, column=1, value=uid)
            ws.cell(row=row, column=2, value=raw_records[uid]['id'])
            ws.cell(row=row, column=3, value=raw_records[uid]['name'].title())
            ws.cell(row=row, column=4, value=f"Week {wk}")
            ws.cell(row=row, column=5, value=to_excel_time(wk_hrs_dec)).number_format = TIME_FMT
            ws.cell(row=row, column=6, value=to_excel_time(wk_target)).number_format  = TIME_FMT
            ws.cell(row=row, column=7, value=f"=MAX(0,E{row}-F{row})").number_format  = TIME_FMT
            ws.cell(row=row, column=8, value=f"=MAX(0,F{row}-E{row})").number_format  = TIME_FMT
            # RawTarget: standard target unadjusted for leave — used only
            # as the salary rate's denominator (see write_consolidated_sheet).
            ws.cell(row=row, column=9, value=to_excel_time(wk_target_raw)).number_format = TIME_FMT

            row_map[uid][wk] = row
            row += 1

    return ws, row_map

def write_summary_sheet(wb, employees_dec, emp_order, raw_records,
                         period_str, year, month, daily_target, part_time_list,
                         pt_daily_target, row_map, holiday_dates, wfh_records):
    ws = wb.create_sheet("Weekly Summary")

    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value     = f"Weekly Attendance Summary | {period_str}"
    c.font      = make_font(True, C_HEADER_FG, 14)
    c.fill      = make_fill(C_HEADER_BG)
    c.alignment = make_align()

    ws.merge_cells("A2:H2")
    note = ws["A2"]
    note.value     = "⚠️  Edit 'Hours Worked' values here (as H:MM, e.g. 8:30) — Consolidated Report updates automatically via formulas."
    note.font      = make_font(False, "7B3F00", 9)
    note.fill      = make_fill(C_HOLIDAY_BG)
    note.alignment = make_align()

    headers = ["ID", "Employee Name", "Week",
               "Hours Worked ✏️", "Target Hours", "Excess", "Shortage", "Status"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = (
            make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()
        )

    row = 5
    for uid in emp_order:
        if uid not in employees_dec or uid not in row_map:
            continue
        current_daily = pt_daily_target if uid in part_time_list else daily_target
        leave_by_week = get_leave_days_by_week(uid, raw_records, year, month,
                                                holiday_dates, wfh_records)
        for wk in sorted(employees_dec[uid].keys()):
            raw_row = row_map[uid].get(wk)
            if raw_row is None:
                continue

            wk_hrs_dec    = sum_week_hours(employees_dec[uid][wk])
            wk_target     = get_effective_week_target(wk, year, month, current_daily, leave_by_week)

            ws.cell(row=row, column=1, value=raw_records[uid]['id'])
            ws.cell(row=row, column=2, value=raw_records[uid]['name'].title())
            ws.cell(row=row, column=3, value=f"Week {wk}")
            ws.cell(row=row, column=4, value=to_excel_time(wk_hrs_dec)).number_format = TIME_FMT
            ws.cell(row=row, column=5, value=to_excel_time(wk_target)).number_format  = TIME_FMT
            ws.cell(row=row, column=6, value=f"=MAX(0,D{row}-E{row})").number_format  = TIME_FMT
            ws.cell(row=row, column=7, value=f"=MAX(0,E{row}-D{row})").number_format  = TIME_FMT
            ws.cell(row=row, column=8,
                    value=f'=IF(D{row}>E{row},"EXCESS",IF(D{row}<E{row},"SHORTAGE","ON TARGET"))')

            row_map[uid][wk] = (raw_row, row)

            if wk_hrs_dec > wk_target:
                fill_c = C_EXCESS_BG
            elif wk_hrs_dec < wk_target:
                fill_c = C_SHORT_BG
            else:
                fill_c = C_ALT_ROW

            for col in range(1, 9):
                c = ws.cell(row=row, column=col)
                c.fill      = make_fill(fill_c)
                c.border    = make_border()
                c.alignment = make_align()
                c.font      = make_font()
                if col == 4:
                    c.font = make_font(bold=True)

            row += 1

    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 14
    for ltr in ['F', 'G', 'H']:
        ws.column_dimensions[ltr].width = 14

    return ws, row_map

# ── CHANGE 3: Net split into Net Hours (number) + Status (label) ──────────────
def write_consolidated_sheet(wb, employees_dec, emp_order, raw_records, period_str,
                              year, month, daily_target, part_time_list, pt_daily_target,
                              holiday_dates, wfh_records, row_map, salary_map=None,
                              ft_payroll_daily_hours=8.3, pt_payroll_daily_hours=8.0):
    ws = wb.create_sheet("Consolidated Report")
    salary_map = salary_map or {}

    headers = [
        "ID", "Employee Name",
        "Total Hours", "Total Target", "Total Excess", "Total Shortage",
        "Net Hours",    # G — number only (positive=excess, negative=shortage)
        "Status",       # H — "Excess" / "Shortage" / "On Target"
        "Days Worked", "Leave Days", "Holidays on Leave", "Sundays", "Holidays",
        "Monthly Salary", "Sal Per Day", "Days Worked (Payroll)", "Gross Salary",
        "Sal Per Hour", "Net Hours (hrs)", "Extra Sal", "Calculated Salary",
        "Net Hours (raw)",   # V — hidden helper: plain number, for math only
    ]
    num_cols = len(headers)
    days_in_month = calendar.monthrange(year, month)[1]

    ws.merge_cells(f"A1:{get_column_letter(num_cols)}1")
    c = ws["A1"]
    c.value     = f"Total Monthly Consolidation | {period_str}"
    c.font      = make_font(True, C_HEADER_FG, 14)
    c.fill      = make_fill(C_HEADER_BG)
    c.alignment = make_align()

    hdr_row = 3
    if holiday_dates:
        ws.merge_cells(f"A2:{get_column_letter(num_cols)}2")
        ws["A2"].value     = "Holidays: " + ", ".join(hd.strftime("%d-%b-%Y") for hd in sorted(holiday_dates))
        ws["A2"].font      = make_font(True, "7B3F00", 9)
        ws["A2"].fill      = make_fill(C_HOLIDAY_BG)
        ws["A2"].alignment = make_align()
        hdr_row = 4

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=hdr_row, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = (
            make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()
        )

    num_sundays  = len(get_month_sundays(year, month))
    num_holidays = len([hd for hd in holiday_dates if hd.year == year and hd.month == month])

    total_summary_rows = sum(len(d) for d in employees_dec.values())
    sum_end_row        = max(5, 5 + total_summary_rows - 1)
    ws_ref             = "'Weekly Summary'"

    data_row = hdr_row + 1

    for uid in emp_order:
        if uid not in employees_dec:
            continue

        emp_name_title = raw_records[uid]['name'].title()
        emp_id         = raw_records[uid]['id']

        name_col     = f"{ws_ref}!$B$5:$B${sum_end_row}"
        hours_col    = f"{ws_ref}!$D$5:$D${sum_end_row}"
        target_col   = f"{ws_ref}!$E$5:$E${sum_end_row}"
        excess_col   = f"{ws_ref}!$F$5:$F${sum_end_row}"
        shortage_col = f"{ws_ref}!$G$5:$G${sum_end_row}"
        crit         = f'"{emp_name_title}"'

        # Values here are day-fractions (hours/24) so cells can be formatted
        # as real elapsed time ([h]:mm = 60 min/hr) instead of base-10
        # decimal. Rounding to 4 dp keeps ~0.35-second precision.
        #
        # Total Hours is capped at each week's target (excess excluded) —
        # actual hours worked minus the excess portion, so a week with
        # overtime doesn't inflate this figure. Total Excess (col E) still
        # shows the excess separately. Total Hours + Total Excess always
        # recovers the true raw hours worked, since Excess is never
        # negative — unlike Net Hours, which can be negative and would
        # double-subtract an already-reflected shortage if added here.
        f_hrs      = (f"=ROUND(SUMIF({name_col},{crit},{hours_col})"
                       f"-MAX(0,SUMIF({name_col},{crit},{excess_col})),4)")
        f_target   = f"=ROUND(SUMIF({name_col},{crit},{target_col}),4)"
        f_excess   = f"=ROUND(MAX(0,SUMIF({name_col},{crit},{excess_col})),4)"
        f_shortage = f"=ROUND(MAX(0,SUMIF({name_col},{crit},{shortage_col})),4)"

        # V: Net Hours as a plain number (Total Excess - Total Shortage,
        # converted from day-fraction to real hours via *24) — hidden, math-
        # only helper. G, H, and S (Net Hours (hrs)) all derive FROM this
        # single cell rather than each recomputing (E-F) themselves, so
        # they can never disagree with each other or with salary math.
        f_net_hours_raw = f"=ROUND((E{data_row}-F{data_row})*24,4)"

        # G and S: Net Hours as text — Excel can't render a negative
        # [h]:mm duration (always shows ####), so build a text label
        # instead: positive net formats normally, negative net gets a "-"
        # prefix on the absolute difference. Divide V back by 24 since
        # TEXT("[h]:mm") expects a day-fraction, not plain hours. S is
        # display-only (identical to G) — Extra Sal (T) uses V, not S,
        # for arithmetic since S is text.
        f_net_text = (
            f'=IF(V{data_row}>=0,TEXT(V{data_row}/24,"[h]:mm"),'
            f'"-"&TEXT(-V{data_row}/24,"[h]:mm"))'
        )

        # H: Status label — based on the same Net Hours figure
        f_status = (
            f'=IF(V{data_row}>0,"Excess",'
            f'IF(V{data_row}<0,"Shortage","On Target"))'
        )

        days_worked       = get_days_worked(uid, raw_records, wfh_records, holiday_dates, year, month)
        leave_days        = get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records)
        holidays_on_leave = get_holidays_on_leave(uid, raw_records, year, month, holiday_dates, wfh_records)

        wk_dict           = employees_dec[uid]
        current_daily     = pt_daily_target if uid in part_time_list else daily_target
        leave_by_week     = get_leave_days_by_week(uid, raw_records, year, month,
                                                     holiday_dates, wfh_records)
        total_hours_dec   = sum(sum(d.values()) for d in wk_dict.values())
        total_target_dec  = sum(get_effective_week_target(wk, year, month, current_daily, leave_by_week)
                                 for wk in wk_dict)
        net               = round(total_hours_dec - total_target_dec, 2)

        monthly_salary = salary_map.get(normalize_id(emp_id))

        if monthly_salary:
            # Days-based payroll formula (matches the existing manual salary
            # sheet), fully formula-driven so it recalculates in Excel if
            # hours or leave are edited afterward:
            #   Sal Per Day (O)            = Monthly Salary / Days in Month
            #   Days Worked, Payroll (P)   = Days in Month - Leave Days (J);
            #     i.e. every paid calendar day (including auto-credited
            #     Sundays/holidays) counts as worked, only real leave
            #     doesn't — unlike the "Days Worked" column (I), which
            #     excludes Sundays/holidays for attendance-tracking purposes.
            #   Gross Salary (Q)           = Sal Per Day * Days Worked, Payroll
            #   Sal Per Hour (R)           = Sal Per Day / standard daily hours
            #   Extra Sal (T)              = Net Hours (raw, V) * Sal Per
            #     Hour — the same figure shown as text in G/S, just the
            #     hidden numeric version, since S is text and can't be
            #     used in arithmetic.
            #   Calculated Salary (U)      = Gross Salary + Extra Sal
            #
            # "Standard daily hours" here is the payroll-specific constant
            # (8.30 FT / 8.00 PT by default) — NOT the attendance daily
            # target (daily_target/pt_daily_target), which is a separate
            # figure used only for Target Hours/Excess/Shortage tracking.
            payroll_daily = pt_payroll_daily_hours if uid in part_time_list else ft_payroll_daily_hours
            f_sal_per_day  = f"=ROUND(N{data_row}/{days_in_month},4)"
            f_days_payroll = f"={days_in_month}-J{data_row}"
            f_gross        = f"=ROUND(O{data_row}*P{data_row},2)"
            f_sal_per_hour = f"=ROUND(O{data_row}/{payroll_daily},4)"
            f_extra_sal    = f"=ROUND(V{data_row}*R{data_row},2)"
            f_calc_salary  = f"=ROUND(Q{data_row}+T{data_row},2)"
        else:
            f_sal_per_day   = None
            f_days_payroll  = None
            f_gross         = None
            f_sal_per_hour  = None
            f_extra_sal     = None
            f_calc_salary   = None

        vals = [
            emp_id,             # A — 1
            emp_name_title,     # B — 2
            f_hrs,              # C — 3
            f_target,           # D — 4
            f_excess,           # E — 5
            f_shortage,         # F — 6
            f_net_text,         # G — 7  Net Hours (text label, e.g. "-8:30")
            f_status,           # H — 8  Status label
            days_worked,        # I — 9
            leave_days,         # J — 10
            holidays_on_leave,  # K — 11
            num_sundays,        # L — 12
            num_holidays,       # M — 13
            monthly_salary,     # N — 14  Monthly Salary (input value)
            f_sal_per_day,      # O — 15  Sal Per Day (formula)
            f_days_payroll,     # P — 16  Days Worked, Payroll (formula)
            f_gross,            # Q — 17  Gross Salary (formula)
            f_sal_per_hour,     # R — 18  Sal Per Hour (formula)
            f_net_text,         # S — 19  Net Hours (hrs) — same text as G
            f_extra_sal,        # T — 20  Extra Sal (formula)
            f_calc_salary,      # U — 21  Calculated Salary (formula)
            f_net_hours_raw,    # V — 22  Net Hours (raw), hidden math helper
        ]

        for col, v in enumerate(vals, 1):
            c = ws.cell(row=data_row, column=col, value=v)
            if col in (3, 4, 5, 6):
                c.number_format = TIME_FMT
            if col in (14, 15, 17, 18, 20, 21):
                c.number_format = "#,##0.00"
            if col == 16:
                c.number_format = "0"
            if col == 22:
                c.number_format = "0.00"
            if col in (7, 8):
                fill_c = C_EXCESS_BG if net > 0 else (C_SHORT_BG if net < 0 else C_ALT_ROW)
            elif col == 11 and isinstance(v, int) and v > 0:
                fill_c = C_HOLIDAY_BG
            elif col == 21 and monthly_salary is None:
                fill_c = C_SHORT_BG
            else:
                fill_c = C_ALT_ROW
            c.fill, c.border, c.alignment, c.font = (
                make_fill(fill_c), make_border(), make_align(), make_font()
            )

        data_row += 1

    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['K'].width = 18
    for ltr in ['C', 'D', 'E', 'F']:
        ws.column_dimensions[ltr].width = 15
    for ltr in ['I', 'J', 'L', 'M']:
        ws.column_dimensions[ltr].width = 13
    for ltr in ['N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U']:
        ws.column_dimensions[ltr].width = 16
    ws.column_dimensions['V'].hidden = True

def write_individual_sheet(wb, uid, week_dict, period_str, year, month,
                           daily_target, is_part_time, pt_daily_target,
                           holiday_dates, wfh_records, raw_records):
    ws_name = (uid[:28]
               .replace(":", "").replace("/", "").replace("*", "")
               .replace("?", "").replace("[", "").replace("]", "").strip())
    ws = wb.create_sheet(ws_name)

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value     = f"Attendance Details | {uid} {'(Part-Time)' if is_part_time else ''}"
    c.font      = make_font(True, C_HEADER_FG, 12)
    c.fill      = make_fill(C_HEADER_BG)
    c.alignment = make_align()

    holiday_day_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    wfh_dict         = wfh_records.get(uid, {})
    punched_days     = set(raw_records[uid]['punches'].keys())
    current_daily    = pt_daily_target if is_part_time else daily_target
    leave_by_week    = get_leave_days_by_week(uid, raw_records, year, month,
                                               holiday_dates, wfh_records)

    row = 3
    for wk in sorted(week_dict.keys()):
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1, value=f"WEEK {wk}")
        c.font, c.fill, c.alignment = make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align()
        row += 1

        for col, h in enumerate(["Date", "Day", "Hours Worked", "Note"], 1):
            c = ws.cell(row=row, column=col, value=h)
            c.font, c.fill, c.alignment, c.border = (
                make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()
            )
        row += 1

        for day, hrs in sorted(week_dict[wk].items()):
            try:
                dt_obj = date(year, month, day)
                d_str  = dt_obj.strftime("%d-%b-%Y")
                d_name = dt_obj.strftime("%A")
            except:
                d_str, d_name = f"Day {day}", ""

            is_holiday    = day in holiday_day_nums
            is_wfh        = day in wfh_dict
            is_auto_sunday = (d_name == "Sunday" and not is_holiday and not is_wfh
                               and day not in punched_days)

            if is_holiday:
                fill_c, note = C_HOLIDAY_BG, f"Holiday (Paid – {decimal_to_hhmm(hrs)} hrs)"
            elif is_wfh:
                info   = wfh_dict[day]
                fill_c = C_WFH_BG
                note   = f"WFH  {info.get('in','?')} → {info.get('out','?')}"
            elif is_auto_sunday:
                fill_c, note = C_HOLIDAY_BG, f"Sunday (Paid – {decimal_to_hhmm(hrs)} hrs)"
            else:
                fill_c, note = C_ALT_ROW, ""

            for col, v in enumerate([d_str, d_name, decimal_to_hhmm(hrs), note], 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill, c.border, c.alignment, c.font = (
                    make_fill(fill_c), make_border(), make_align(), make_font()
                )
            row += 1

        wk_target    = get_effective_week_target(wk, year, month, current_daily, leave_by_week)
        wk_hrs_dec   = sum_week_hours(week_dict[wk])
        excess       = max(0.0, wk_hrs_dec - wk_target)
        shortage     = max(0.0, wk_target  - wk_hrs_dec)
        summary_fill = make_fill(C_NEUTRAL_BG)

        for label, val in [
            ("Total Worked (Week)", decimal_to_hhmm(wk_hrs_dec)),
            ("Target (Week)",       decimal_to_hhmm(wk_target)),
            ("Excess Hours",        decimal_to_hhmm(excess)),
            ("Shortage Hours",      decimal_to_hhmm(shortage)),
        ]:
            ws.merge_cells(f"A{row}:C{row}")
            for col, v in enumerate([label, None, None, val], 1):
                if col in (2, 3): continue
                c = ws.cell(row=row, column=col, value=v)
                c.font, c.fill, c.border, c.alignment = (
                    make_font(True), summary_fill, make_border(), make_align()
                )
            row += 1
        row += 1

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 26

def write_wfh_sheet(wb, emp_order, raw_records, wfh_records, year, month, period_str):
    ws = wb.create_sheet("WFH Log")

    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value, c.font, c.fill, c.alignment = (
        f"Work From Home Log | {period_str}",
        make_font(True, C_HEADER_FG, 14), make_fill(C_HEADER_BG), make_align()
    )

    for col, h in enumerate(["ID", "Employee Name", "Date", "Time In → Out", "Hours"], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = (
            make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()
        )

    row = 4
    for uid in emp_order:
        wfh_dict = wfh_records.get(uid, {})
        if not wfh_dict:
            continue
        for day in sorted(wfh_dict.keys()):
            info = wfh_dict[day]
            try:
                d_str = date(year, month, day).strftime("%d-%b-%Y (%a)")
            except:
                d_str = f"Day {day}"
            for col, v in enumerate([
                raw_records[uid]['id'],
                raw_records[uid]['name'].title(),
                d_str,
                f"{info.get('in','?')} → {info.get('out','?')}",
                decimal_to_hhmm(info.get('hours', 0.0))
            ], 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill, c.border, c.alignment, c.font = (
                    make_fill(C_WFH_BG), make_border(), make_align(), make_font()
                )
            row += 1

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 12

def generate_report(employees_dec, emp_order, raw_records, period_str,
                    year, month, daily_target, part_time_list, pt_daily_target,
                    holiday_dates, wfh_records, salary_map=None,
                    ft_payroll_daily_hours=8.3, pt_payroll_daily_hours=8.0):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _, row_map = write_raw_data_sheet(wb, employees_dec, emp_order, raw_records,
                                      year, month, daily_target, part_time_list,
                                      pt_daily_target, period_str, holiday_dates, wfh_records)

    _, row_map = write_summary_sheet(wb, employees_dec, emp_order, raw_records,
                                     period_str, year, month, daily_target,
                                     part_time_list, pt_daily_target, row_map,
                                     holiday_dates, wfh_records)

    write_consolidated_sheet(wb, employees_dec, emp_order, raw_records, period_str,
                              year, month, daily_target, part_time_list, pt_daily_target,
                              holiday_dates, wfh_records, row_map, salary_map,
                              ft_payroll_daily_hours, pt_payroll_daily_hours)

    write_wfh_sheet(wb, emp_order, raw_records, wfh_records, year, month, period_str)

    for uid in emp_order:
        if uid in employees_dec:
            write_individual_sheet(wb, uid, employees_dec[uid], period_str, year, month,
                                   daily_target, uid in part_time_list, pt_daily_target,
                                   holiday_dates, wfh_records, raw_records)

    sheet_order = ["Weekly Summary", "Consolidated Report", "WFH Log", RAW_SHEET]
    for uid in emp_order:
        if uid in employees_dec:
            ws_name = (uid[:28]
                       .replace(":", "").replace("/", "").replace("*", "")
                       .replace("?", "").replace("[", "").replace("]", "").strip())
            sheet_order.append(ws_name)

    existing  = [s.title for s in wb.worksheets]
    ordered   = [s for s in sheet_order if s in existing]
    remaining = [s for s in existing  if s not in ordered]
    for i, name in enumerate(ordered + remaining):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def main():
    st.set_page_config(page_title="Attendance Processor", layout="wide")
    st.title("🕐 Attendance Processor")

    uploaded = st.file_uploader("📂 Upload Filled Attendance XLSX (attendance + logs + salary sheets)", type=["xlsx"])
    if not uploaded:
        return

    data = load_filled_workbook(uploaded.getvalue())
    if data is None:
        st.error("No attendance sheet found (expected a sheet with 'In Time' / 'Out Time' rows).")
        return
    if not data['has_logs']:
        st.warning("⚠️ No logs sheet ('No :' / 'Name :' layout) found — IDs can't be matched.")
    if not data['has_salary']:
        st.warning("⚠️ No salary sheet ('Name of the Staff' / 'Basic Salary') found — salary preview disabled.")

    year, month = data['year'], data['month']
    if year is None:
        st.error("Couldn't read the month from the attendance sheet.")
        return
    last_day   = calendar.monthrange(year, month)[1]
    period_str = data['period_str'] or f"{year}/{month:02d}/01 ~ {month:02d}/{last_day:02d}"

    blocks = [b for b in data['blocks'] if b['name'] and b['punches']]

    # Part-time = daily target below the full-time one in the sheet (8:30 FT / 8:00 PT).
    targets      = sorted({round(b['target'], 4) for b in blocks if b['target']})
    ft_sheet_tgt = targets[-1] if targets else 8.5
    pt_sheet_tgt = targets[0] if len(targets) > 1 else 8.0
    is_pt_block  = lambda b: bool(b['target']) and b['target'] < ft_sheet_tgt - 1e-6

    # Name → ID (logs sheet) and name → salary row (salary sheet), best guesses
    # first; anything under the threshold is left blank for the user to pick.
    NONE_OPT   = "— none —"
    id_opts    = [f"{i} – {n}" for i, n in data['log_ids']]
    sal_opts   = [f"{n} – {int(s):,}" for n, s in data['salaries']]
    id_guess   = match_names([b['name'] for b in blocks], [n for _, n in data['log_ids']])
    sal_guess  = match_names([b['name'] for b in blocks], [n for n, _ in data['salaries']])

    file_sig = f"{uploaded.name}:{uploaded.size}"
    map_rows = [{
        "Employee":     b['name'],
        "Part-time":    is_pt_block(b),
        "Logs ID":      id_opts[id_guess[k]] if id_guess[k] is not None else NONE_OPT,
        "Salary sheet": sal_opts[sal_guess[k]] if sal_guess[k] is not None else NONE_OPT,
        "Salary override": 0.0,
    } for k, b in enumerate(blocks)]

    st.header("🔗 Match Employees")
    n_review = sum(1 for r in map_rows if NONE_OPT in (r["Logs ID"], r["Salary sheet"]))
    st.caption(
        "Names come from the attendance sheet. The ID is taken from the logs sheet and the "
        "salary from the salary sheet by closest name. Check the rows below and fix any "
        "wrong or blank match — salary uses the override if you enter one."
    )
    if n_review:
        st.warning(f"⚠️ {n_review} employee(s) have a blank ID or salary match — review them below.")
    edited = st.data_editor(
        pd.DataFrame(map_rows),
        key=f"map_editor_{file_sig}",
        hide_index=True, use_container_width=True,
        disabled=["Employee", "Part-time"],
        column_config={
            "Logs ID":      st.column_config.SelectboxColumn(options=[NONE_OPT] + id_opts, required=True),
            "Salary sheet": st.column_config.SelectboxColumn(options=[NONE_OPT] + sal_opts, required=True),
            "Salary override": st.column_config.NumberColumn(min_value=0.0, step=500.0),
        },
    )

    raw_records, emp_order, part_time_list, salary_by_uid = {}, [], [], {}
    for k, b in enumerate(blocks):
        row    = edited.iloc[k]
        id_pick, sal_pick = row["Logs ID"], row["Salary sheet"]
        if id_pick in id_opts:
            emp_no = normalize_id(data['log_ids'][id_opts.index(id_pick)][0])
        else:
            emp_no = f"NA{k + 1}"          # unique placeholder so records never collide
        uid = f"{b['name'].title()} (ID: {emp_no})"
        if uid in raw_records:
            uid = f"{uid} #{k + 1}"
        raw_records[uid] = {'name': b['name'], 'id': emp_no, 'punches': b['punches']}
        emp_order.append(uid)
        if is_pt_block(b):
            part_time_list.append(uid)
        override = float(row["Salary override"] or 0)
        if override > 0:
            salary_by_uid[uid] = override
        elif sal_pick in sal_opts:
            salary_by_uid[uid] = data['salaries'][sal_opts.index(sal_pick)][1]
    active_employees = emp_order

    for key, default in [
        ('holiday_dates',  []),
        ('wfh_records',    {}),
        ('fixes',          {}),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    # New workbook → start from what the sheet says (yellow weekday = office
    # holiday) instead of carrying over last month's manual entries.
    if st.session_state.get('loaded_sig') != file_sig:
        st.session_state.loaded_sig    = file_sig
        st.session_state.holiday_dates = [date(year, month, d) for d in data['holiday_days']]
        st.session_state.wfh_records   = {}
        st.session_state.fixes         = {}

    with st.sidebar:
        st.header("⚙️ Settings")
        target_weekly = st.number_input(
            "Full-Time Weekly Target (hrs, 7-day week)",
            min_value=1.0, value=round(ft_sheet_tgt * 7, 2), step=0.5,
            key=f"ft_weekly_{file_sig}",
            help="Defaulted from the daily target row in the attendance sheet."
        )
        daily_target = round(target_weekly / 7, 10)

        st.divider()
        st.subheader("🕑 Part-Time Settings")
        pt_daily_target = st.number_input(
            "Part-Time Daily Target (hrs)", min_value=0.5, value=float(pt_sheet_tgt), step=0.5,
            key=f"pt_daily_{file_sig}",
            help="Defaulted from the daily target row in the attendance sheet."
        )
        st.caption(
            "Part-time is read from the sheet: anyone whose daily target is below the "
            "full-time target. Change it in the sheet to change it here."
        )
        if part_time_list:
            st.caption(
                "**Part-time (" + str(len(part_time_list)) + "):** "
                + ", ".join(raw_records[u]['name'].title() for u in part_time_list)
            )
        else:
            st.caption("No part-time employees found in the sheet.")

        st.divider()
        st.subheader("🏖️ Office Holidays")
        st.caption("Credited as paid holiday (incl. Sundays) at the hours below — not the attendance daily target above; not counted as a working day.")

        hc1, hc2 = st.columns(2)
        ft_holiday_credit_hours = hc1.number_input(
            "FT Sunday/Holiday Credit (hrs)", min_value=0.1, value=8.5, step=0.05,
            help="Entered as decimal hours, e.g. 8.5 = 8 hours 30 minutes (8:30)."
        )
        pt_holiday_credit_hours = hc2.number_input(
            "PT Sunday/Holiday Credit (hrs)", min_value=0.1, value=8.0, step=0.05
        )

        new_holiday = st.date_input(
            "Pick holiday date",
            value=date(year, month, 1),
            min_value=date(year, month, 1),
            max_value=date(year, month, calendar.monthrange(year, month)[1]),
            key="holiday_picker"
        )
        if st.button("➕ Add Holiday"):
            if new_holiday not in st.session_state.holiday_dates:
                st.session_state.holiday_dates.append(new_holiday)
                st.success(f"Added {new_holiday.strftime('%d-%b-%Y')}")
            else:
                st.warning("Already added.")

        if st.session_state.holiday_dates:
            st.write("**Holidays:**")
            for hd in sorted(st.session_state.holiday_dates):
                c1, c2 = st.columns([3, 1])
                c1.write(hd.strftime("%d-%b-%Y (%a)"))
                if c2.button("✕", key=f"del_hol_{hd}"):
                    st.session_state.holiday_dates.remove(hd)
                    st.rerun()

        st.divider()
        st.subheader("🏠 Work From Home")
        st.caption("Set the employee, date and exact hours worked from home.")

        wfh_emp = st.selectbox(
            "Employee",
            options=active_employees,
            format_func=lambda x: raw_records[x]['name'].title(),
            key="wfh_emp_select"
        ) if active_employees else None

        wfh_date = st.date_input(
            "WFH Date",
            value=date(year, month, 1),
            min_value=date(year, month, 1),
            max_value=date(year, month, calendar.monthrange(year, month)[1]),
            key="wfh_date_picker"
        )

        st.write("**Hours worked from home**")
        ic, oc = st.columns(2)
        with ic:
            st.caption("🟢 Time In")
            wfh_in_h = st.number_input("Hour",   0, 23, 9,  key="wfh_in_h")
            wfh_in_m = st.selectbox("Min", [0, 15, 30, 45],
                                     format_func=lambda x: f"{x:02d}", key="wfh_in_m")
        with oc:
            st.caption("🔴 Time Out")
            wfh_out_h = st.number_input("Hour",   0, 23, 18, key="wfh_out_h")
            wfh_out_m = st.selectbox("Min", [0, 15, 30, 45],
                                      format_func=lambda x: f"{x:02d}", key="wfh_out_m")

        wfh_in_str  = f"{int(wfh_in_h):02d}:{int(wfh_in_m):02d}"
        wfh_out_str = f"{int(wfh_out_h):02d}:{int(wfh_out_m):02d}"
        wfh_hrs, _  = compute_hours_from_pair(wfh_in_str, wfh_out_str)

        if wfh_hrs > 0:
            st.info(f"⏱ {wfh_in_str} → {wfh_out_str} = **{decimal_to_hhmm(wfh_hrs)} hrs**")
        else:
            st.warning("⚠️ Out time must be after In time.")

        if st.button("➕ Add WFH Day", disabled=(wfh_hrs <= 0)):
            if wfh_emp:
                if wfh_emp not in st.session_state.wfh_records:
                    st.session_state.wfh_records[wfh_emp] = {}
                st.session_state.wfh_records[wfh_emp][wfh_date.day] = {
                    'in': wfh_in_str, 'out': wfh_out_str, 'hours': wfh_hrs
                }
                st.success(
                    f"✅ {raw_records[wfh_emp]['name'].title()} | "
                    f"{wfh_date.strftime('%d-%b-%Y')} | "
                    f"{wfh_in_str}→{wfh_out_str} | {decimal_to_hhmm(wfh_hrs)} hrs"
                )

        any_wfh = any(v for v in st.session_state.wfh_records.values())
        if any_wfh:
            st.write("**Current WFH log:**")
            for uid in emp_order:
                wd = st.session_state.wfh_records.get(uid, {})
                if not wd:
                    continue
                st.markdown(f"**{raw_records[uid]['name'].title()}**")
                for d in sorted(wd.keys()):
                    info = wd[d]
                    try:
                        d_lbl = date(year, month, d).strftime("%d-%b (%a)")
                    except:
                        d_lbl = f"Day {d}"
                    cx, cy = st.columns([4, 1])
                    cx.write(
                        f"{d_lbl}  {info.get('in','?')}→{info.get('out','?')}  "
                        f"({decimal_to_hhmm(info.get('hours', 0))} hrs)"
                    )
                    if cy.button("✕", key=f"del_wfh_{uid}_{d}"):
                        del st.session_state.wfh_records[uid][d]
                        st.rerun()

        st.divider()
        st.subheader("💰 Salary Settings")
        st.caption(
            "Monthly salary is the Basic Salary from the salary sheet, matched by name. "
            "Days-based payroll formula: Sal Per Day = Salary ÷ Days in Month; "
            "Gross Salary = Sal Per Day × (Days in Month − Leave Days); "
            "Extra Sal = Net Hours (Excess − Shortage) × Sal Per Hour; "
            "Calculated Salary = Gross Salary + Extra Sal."
        )

        sc1, sc2 = st.columns(2)
        ft_payroll_daily_hours = sc1.number_input(
            "FT Standard Daily Hours (salary rate)", min_value=0.1, value=8.3, step=0.05,
            help="Used only for the payroll Sal Per Hour calculation — "
                 "separate from the Full-Time Weekly Target above (which "
                 "drives Target Hours/Excess/Shortage) and from the Sunday/"
                 "Holiday Credit hours set under Office Holidays."
        )
        pt_payroll_daily_hours = sc2.number_input(
            "PT Standard Daily Hours (salary rate)", min_value=0.1, value=8.0, step=0.05,
            help="Same as above, for part-time employees."
        )


    holiday_dates = st.session_state.holiday_dates
    wfh_records   = st.session_state.wfh_records
    salary_map    = {raw_records[u]['id']: s for u, s in salary_by_uid.items()}
    dup_ids = {i for i in (raw_records[u]['id'] for u in active_employees)
               if sum(raw_records[u]['id'] == i for u in active_employees) > 1}
    if dup_ids:
        st.error("⚠️ The same logs ID is assigned to more than one employee: "
                 + ", ".join(sorted(dup_ids)) + ". Fix it in the match table above.")

    st.header("🔧 Fix Missing Punches")
    any_missing = False
    for uid in active_employees:
        p_dict    = raw_records[uid]['punches']
        emp_fixes = st.session_state.fixes.get(uid, {})
        is_pt = uid in part_time_list
        for day, p in sorted(p_dict.items()):
            if len(p) == 1:
                any_missing = True
                c1, c2, c3, c4 = st.columns([2, 1, 3, 2])
                c1.markdown(f"**{uid}**")
                c2.write(f"Day {day} ({'PT' if is_pt else 'FT'})")
                default_in, default_out = infer_missing_punch(p[0], is_pt)
                if default_out == p[0]:
                    c3.warning(f"Out: {p[0]} (In missing)")
                    f_in = c4.text_input("Set In (HH:MM)", value=default_in, key=f"{uid}_{day}_in")
                    try:
                        datetime.strptime(f_in, "%H:%M")
                        emp_fixes[day] = {'in': f_in, 'out': p[0]}
                    except:
                        c4.error("Use HH:MM")
                else:
                    c3.warning(f"In: {p[0]} (Out missing)")
                    f_out = c4.text_input("Set Out (HH:MM)", value=default_out, key=f"{uid}_{day}_out")
                    try:
                        datetime.strptime(f_out, "%H:%M")
                        emp_fixes[day] = {'in': p[0], 'out': f_out}
                    except:
                        c4.error("Use HH:MM")
        if emp_fixes:
            st.session_state.fixes[uid] = emp_fixes

    if not any_missing:
        st.info("✅ No missing punches detected.")

    employees_dec = build_employees_dec(
        active_employees, raw_records, st.session_state.fixes, wfh_records, year, month,
        holiday_dates, part_time_list, pt_daily_target, daily_target,
        ft_holiday_credit_hours, pt_holiday_credit_hours
    )

    st.header("📊 Attendance Summary Preview")
    num_sundays  = len(get_month_sundays(year, month))
    num_holidays = len([hd for hd in holiday_dates if hd.year == year and hd.month == month])

    preview = []
    for uid in active_employees:
        if uid not in employees_dec:
            continue
        wd           = wfh_records.get(uid, {})
        wfh_count    = len(wd)
        total_wfh_h  = sum(v.get('hours', 0.0) for v in wd.values())
        days_worked  = get_days_worked(uid, raw_records, wfh_records, holiday_dates, year, month)
        leave_days   = get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records)
        hol_on_leave = get_holidays_on_leave(uid, raw_records, year, month, holiday_dates, wfh_records)

        week_dict     = employees_dec[uid]
        current_daily = pt_daily_target if uid in part_time_list else daily_target
        leave_by_week = get_leave_days_by_week(uid, raw_records, year, month,
                                                holiday_dates, wfh_records)
        for wk in sorted(week_dict.keys()):
            wk_target     = get_effective_week_target(wk, year, month, current_daily, leave_by_week)
            wk_hrs        = sum_week_hours(week_dict[wk])
            net           = round(wk_hrs - wk_target, 2)
            preview.append({
                "Employee":          raw_records[uid]['name'].title(),
                "Week":              f"Week {wk}",
                "Hrs Worked":        decimal_to_hhmm(wk_hrs),
                "Target":            decimal_to_hhmm(wk_target),
                "Excess":            decimal_to_hhmm(max(0, wk_hrs - wk_target)),
                "Shortage":          decimal_to_hhmm(max(0, wk_target - wk_hrs)),
                "Net Hours":         net,
                "Status":            "Excess" if net > 0 else ("Shortage" if net < 0 else "On Target"),
                "WFH Days":          wfh_count,
                "WFH Hrs":           decimal_to_hhmm(total_wfh_h),
                "Leave Days":        leave_days,
                "Holidays on Leave": hol_on_leave,
                "Sundays":           num_sundays,
                "Holidays":          num_holidays,
            })

    if preview:
        st.dataframe(preview, use_container_width=True)
        st.info(
            "💡 **How editing works in Excel:** Edit the "
            "**'Hours Worked ✏️'** column (col D) in the **Weekly Summary** sheet — "
            "the **Consolidated Report** updates automatically via Excel SUMIF formulas."
        )
    else:
        st.info("No data to preview yet.")

    if salary_map:
        st.header("💰 Salary Preview")
        days_in_month = calendar.monthrange(year, month)[1]
        salary_preview = []
        for uid in active_employees:
            if uid not in employees_dec:
                continue
            emp_id = raw_records[uid]['id']
            monthly_salary = salary_map.get(emp_id)
            if not monthly_salary:
                continue
            current_daily    = pt_daily_target if uid in part_time_list else daily_target
            week_dict        = employees_dec[uid]
            leave_by_week    = get_leave_days_by_week(uid, raw_records, year, month,
                                                        holiday_dates, wfh_records)
            leave_days       = get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records)
            total_hours_dec  = sum(sum(d.values()) for d in week_dict.values())
            total_target_dec = sum(get_effective_week_target(wk, year, month, current_daily, leave_by_week)
                                    for wk in week_dict)
            total_excess_dec = sum(max(0.0, sum(d.values())
                                    - get_effective_week_target(wk, year, month, current_daily, leave_by_week))
                                    for wk, d in week_dict.items())
            capped_hours_dec = total_hours_dec - total_excess_dec
            net              = round(total_hours_dec - total_target_dec, 2)

            # Days-based payroll formula (matches the manual salary sheet) —
            # see write_consolidated_sheet for the full explanation. Uses the
            # payroll-specific standard daily hours (FT/PT), not the
            # attendance daily target used for Target Hours/Excess/Shortage.
            # The hour-level pay adjustment uses Net Hours (Excess - Shortage,
            # same figure as the Net Hours column), not a separate days-based
            # "extra time" comparison.
            payroll_daily  = pt_payroll_daily_hours if uid in part_time_list else ft_payroll_daily_hours
            sal_per_day    = monthly_salary / days_in_month
            days_payroll   = days_in_month - leave_days
            gross_salary   = round(sal_per_day * days_payroll, 2)
            sal_per_hour   = sal_per_day / payroll_daily
            extra_sal      = round(net * sal_per_hour, 2)
            calc_salary    = round(gross_salary + extra_sal, 2)

            salary_preview.append({
                "Employee":              raw_records[uid]['name'].title(),
                "Total Hours":           decimal_to_hhmm(capped_hours_dec),
                "Total Target":          decimal_to_hhmm(total_target_dec),
                "Excess":                decimal_to_hhmm(total_excess_dec),
                "Net Hours":             net,
                "Monthly Salary":        monthly_salary,
                "Days Worked (Payroll)": days_payroll,
                "Gross Salary":          gross_salary,
                "Extra Sal":             extra_sal,
                "Calculated Salary":     calc_salary,
            })
        if salary_preview:
            st.dataframe(salary_preview, use_container_width=True)
        missing = [f"{raw_records[uid]['name'].title()} (ID: {raw_records[uid]['id']})"
                   for uid in active_employees
                   if uid in employees_dec and not salary_map.get(raw_records[uid]['id'])]
        if missing:
            st.warning("⚠️ No salary set for: " + ", ".join(missing))

    st.header("📥 Download Final Report")
    if st.button("Generate Excel Report", type="primary"):
        buf = generate_report(
            employees_dec, active_employees, raw_records, period_str,
            year, month, daily_target, part_time_list, pt_daily_target,
            holiday_dates, wfh_records, salary_map,
            ft_payroll_daily_hours, pt_payroll_daily_hours
        )
        st.download_button(
            "⬇️ Download attendance_report.xlsx",
            buf, "attendance_report.xlsx"
        )

if __name__ == "__main__":
    main()
