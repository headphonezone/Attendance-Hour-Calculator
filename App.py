import io
from datetime import datetime, date
from collections import defaultdict
import calendar

import streamlit as st
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
                if date(year, month, d_int).weekday() != 6:
                    wd += 1
        except ValueError:
            break
    return round(wd * daily_target, 2)

def get_month_sundays(year, month):
    total = calendar.monthrange(year, month)[1]
    return [d for d in range(1, total + 1) if date(year, month, d).weekday() == 6]

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
            emp_no    = str(row[2]).strip()  if row[2]  else 'Unknown'
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
                        if date(year, month, day_num).weekday() == 6: continue
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

# ── CHANGE 1 + 2: Skip relieved employees + inject paid holiday 8.5 hrs ──────
def build_employees_dec(emp_order, raw_records, fixes, wfh_records, year, month,
                         holiday_dates):
    employees_dec    = {}
    holiday_day_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)

    for uid in emp_order:
        p_dict  = raw_records[uid]['punches']
        has_wfh = bool(wfh_records.get(uid, {}))

        # Skip relieved employees — no punches and no WFH for the entire month
        if not p_dict and not has_wfh:
            continue

        f_dict    = fixes.get(uid, {})
        week_data = defaultdict(dict)

        for day, p in p_dict.items():
            if day in f_dict:
                dec, _ = compute_hours_from_pair(f_dict[day]['in'], f_dict[day]['out'])
            elif len(p) >= 2:
                dec = sum(compute_hours_from_pair(p[i], p[i+1])[0] for i in range(0, len(p)-1, 2))
            elif len(p) == 1:
                h   = int(p[0].split(':')[0]) if ':' in p[0] else 0
                dec, _ = (compute_hours_from_pair("09:30", p[0]) if h >= 12
                          else compute_hours_from_pair(p[0], "18:00"))
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

        # Inject paid holiday 8.5 hrs — credited to hours but NOT a working day.
        # Only fills days with no punch/WFH entry.
        for hday in holiday_day_nums:
            try:
                if date(year, month, hday).weekday() == 6:
                    continue   # skip Sunday holidays
            except:
                continue
            wk = get_week_number(hday, year, month)
            if hday not in week_data.get(wk, {}):
                week_data[wk][hday] = 8.5   # 8 hrs 30 min paid holiday

        if week_data:
            employees_dec[uid] = dict(week_data)
    return employees_dec

def get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records):
    total_days   = calendar.monthrange(year, month)[1]
    punched_days = set(raw_records[uid]['punches'].keys())
    sundays      = set(get_month_sundays(year, month))
    holiday_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    wfh_days     = set(wfh_records.get(uid, {}).keys())
    leave = 0
    for d in range(1, total_days + 1):
        if d in sundays or d in holiday_nums or d in wfh_days: continue
        if d not in punched_days:
            leave += 1
    return leave

def get_holidays_on_leave(uid, raw_records, year, month, holiday_dates, wfh_records):
    punched_days = set(raw_records[uid]['punches'].keys())
    sundays      = set(get_month_sundays(year, month))
    wfh_days     = set(wfh_records.get(uid, {}).keys())
    count = 0
    for hd in holiday_dates:
        if hd.year != year or hd.month != month:
            continue
        d = hd.day
        if d in sundays:    continue
        if d in wfh_days:   continue
        if d not in punched_days:
            count += 1
    return count

# ── CHANGE 2b: get_days_worked — holidays NOT counted as working days ─────────
def get_days_worked(uid, employees_dec, wfh_records, holiday_dates, year, month):
    holiday_day_nums = set(hd.day for hd in holiday_dates if hd.year == year and hd.month == month)
    punched = set()
    if uid in employees_dec:
        for wk_data in employees_dec[uid].values():
            for day, hrs in wk_data.items():
                if hrs > 0 and day not in holiday_day_nums:
                    punched.add(day)
    wfh_non_holiday = {d for d in wfh_records.get(uid, {}) if d not in holiday_day_nums}
    return len(punched | wfh_non_holiday)

def sum_week_hours(day_dict):
    return sum(day_dict.values())

RAW_SHEET          = "_RawData"
RAW_DATA_START_ROW = 2

def write_raw_data_sheet(wb, employees_dec, emp_order, raw_records,
                          year, month, daily_target, part_time_list, pt_daily_target,
                          period_str):
    ws = wb.create_sheet(RAW_SHEET)
    ws.sheet_state = 'hidden'

    for col, h in enumerate(["UID_KEY", "ID", "Name", "Week",
                              "HoursWorked", "Target", "Excess", "Shortage"], 1):
        ws.cell(row=1, column=col, value=h)

    row     = RAW_DATA_START_ROW
    row_map = {}

    for uid in emp_order:
        if uid not in employees_dec:
            continue
        current_daily = pt_daily_target if uid in part_time_list else daily_target
        week_dict     = employees_dec[uid]
        row_map[uid]  = {}

        for wk in sorted(week_dict.keys()):
            wk_target  = get_week_target(wk, year, month, current_daily)
            wk_hrs_dec = sum_week_hours(week_dict[wk])

            ws.cell(row=row, column=1, value=uid)
            ws.cell(row=row, column=2, value=raw_records[uid]['id'])
            ws.cell(row=row, column=3, value=raw_records[uid]['name'].title())
            ws.cell(row=row, column=4, value=f"Week {wk}")
            ws.cell(row=row, column=5, value=round(wk_hrs_dec, 2))
            ws.cell(row=row, column=6, value=round(wk_target,  2))
            ws.cell(row=row, column=7, value=f"=MAX(0,E{row}-F{row})")
            ws.cell(row=row, column=8, value=f"=MAX(0,F{row}-E{row})")

            row_map[uid][wk] = row
            row += 1

    return ws, row_map

def write_summary_sheet(wb, employees_dec, emp_order, raw_records,
                         period_str, year, month, daily_target, part_time_list,
                         pt_daily_target, row_map):
    ws = wb.create_sheet("Weekly Summary")

    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value     = f"Weekly Attendance Summary | {period_str}"
    c.font      = make_font(True, C_HEADER_FG, 14)
    c.fill      = make_fill(C_HEADER_BG)
    c.alignment = make_align()

    ws.merge_cells("A2:H2")
    note = ws["A2"]
    note.value     = "⚠️  Edit 'Hours Worked' values here — Consolidated Report updates automatically via formulas."
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
        for wk in sorted(employees_dec[uid].keys()):
            raw_row = row_map[uid].get(wk)
            if raw_row is None:
                continue

            wk_hrs_dec    = sum_week_hours(employees_dec[uid][wk])
            current_daily = pt_daily_target if uid in part_time_list else daily_target
            wk_target     = get_week_target(wk, year, month, current_daily)

            ws.cell(row=row, column=1, value=raw_records[uid]['id'])
            ws.cell(row=row, column=2, value=raw_records[uid]['name'].title())
            ws.cell(row=row, column=3, value=f"Week {wk}")
            ws.cell(row=row, column=4, value=round(wk_hrs_dec, 2))
            ws.cell(row=row, column=5, value=round(wk_target, 2))
            ws.cell(row=row, column=6, value=f"=MAX(0,D{row}-E{row})")
            ws.cell(row=row, column=7, value=f"=MAX(0,E{row}-D{row})")
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
                              holiday_dates, wfh_records, row_map):
    ws = wb.create_sheet("Consolidated Report")

    headers = [
        "ID", "Employee Name",
        "Total Hours", "Total Target", "Total Excess", "Total Shortage",
        "Net Hours",    # G — number only (positive=excess, negative=shortage)
        "Status",       # H — "Excess" / "Shortage" / "On Target"
        "Days Worked", "Leave Days", "Holidays on Leave", "Sundays", "Holidays"
    ]
    num_cols = len(headers)

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

        f_hrs      = f"=ROUND(SUMIF({name_col},{crit},{hours_col}),2)"
        f_target   = f"=ROUND(SUMIF({name_col},{crit},{target_col}),2)"
        f_excess   = f"=ROUND(MAX(0,SUMIF({name_col},{crit},{excess_col})),2)"
        f_shortage = f"=ROUND(MAX(0,SUMIF({name_col},{crit},{shortage_col})),2)"

        # G: Net Hours — pure number
        e_col = get_column_letter(5)
        f_col = get_column_letter(6)
        g_col = get_column_letter(7)
        f_net_num = f"=ROUND({e_col}{data_row}-{f_col}{data_row},2)"

        # H: Status label only
        f_status = (
            f'=IF({g_col}{data_row}>0,"Excess",'
            f'IF({g_col}{data_row}<0,"Shortage","On Target"))'
        )

        days_worked       = get_days_worked(uid, employees_dec, wfh_records, holiday_dates, year, month)
        leave_days        = get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records)
        holidays_on_leave = get_holidays_on_leave(uid, raw_records, year, month, holiday_dates, wfh_records)

        vals = [
            emp_id,             # A — 1
            emp_name_title,     # B — 2
            f_hrs,              # C — 3
            f_target,           # D — 4
            f_excess,           # E — 5
            f_shortage,         # F — 6
            f_net_num,          # G — 7  Net Hours (number)
            f_status,           # H — 8  Status label
            days_worked,        # I — 9
            leave_days,         # J — 10
            holidays_on_leave,  # K — 11
            num_sundays,        # L — 12
            num_holidays,       # M — 13
        ]

        wk_dict   = employees_dec[uid]
        total_exc = sum(max(0, sum(d.values()) - get_week_target(wk, year, month,
                         pt_daily_target if uid in part_time_list else daily_target))
                         for wk, d in wk_dict.items())
        total_sht = sum(max(0, get_week_target(wk, year, month,
                         pt_daily_target if uid in part_time_list else daily_target) - sum(d.values()))
                         for wk, d in wk_dict.items())
        net = round(total_exc - total_sht, 2)

        for col, v in enumerate(vals, 1):
            c = ws.cell(row=data_row, column=col, value=v)
            if col in (7, 8):
                fill_c = C_EXCESS_BG if net > 0 else (C_SHORT_BG if net < 0 else C_ALT_ROW)
            elif col == 11 and isinstance(v, int) and v > 0:
                fill_c = C_HOLIDAY_BG
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

def write_individual_sheet(wb, uid, week_dict, period_str, year, month,
                           daily_target, is_part_time, pt_daily_target,
                           holiday_dates, wfh_records):
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
    current_daily    = pt_daily_target if is_part_time else daily_target

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

            is_holiday = day in holiday_day_nums
            is_wfh     = day in wfh_dict

            if is_holiday:
                fill_c, note = C_HOLIDAY_BG, "Holiday (Paid – 8.30 hrs)"
            elif is_wfh:
                info   = wfh_dict[day]
                fill_c = C_WFH_BG
                note   = f"WFH  {info.get('in','?')} → {info.get('out','?')}"
            else:
                fill_c, note = C_ALT_ROW, ""

            for col, v in enumerate([d_str, d_name, decimal_to_hhmm(hrs), note], 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill, c.border, c.alignment, c.font = (
                    make_fill(fill_c), make_border(), make_align(), make_font()
                )
            row += 1

        wk_target    = get_week_target(wk, year, month, current_daily)
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
                    holiday_dates, wfh_records):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _, row_map = write_raw_data_sheet(wb, employees_dec, emp_order, raw_records,
                                      year, month, daily_target, part_time_list,
                                      pt_daily_target, period_str)

    _, row_map = write_summary_sheet(wb, employees_dec, emp_order, raw_records,
                                     period_str, year, month, daily_target,
                                     part_time_list, pt_daily_target, row_map)

    write_consolidated_sheet(wb, employees_dec, emp_order, raw_records, period_str,
                              year, month, daily_target, part_time_list, pt_daily_target,
                              holiday_dates, wfh_records, row_map)

    write_wfh_sheet(wb, emp_order, raw_records, wfh_records, year, month, period_str)

    for uid in emp_order:
        if uid in employees_dec:
            write_individual_sheet(wb, uid, employees_dec[uid], period_str, year, month,
                                   daily_target, uid in part_time_list, pt_daily_target,
                                   holiday_dates, wfh_records)

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

    uploaded = st.file_uploader("📂 Upload Attendance XLSX", type=["xlsx"])
    if not uploaded:
        return

    wb_in     = openpyxl.load_workbook(uploaded, read_only=True)
    log_sheet = next((wb_in[n] for n in wb_in.sheetnames if n.lower() == 'logs'), None)
    if not log_sheet:
        st.error("No 'Logs' sheet found.")
        return

    raw_records, emp_order, period_str, year, month = parse_logs_sheet(log_sheet)
    active_employees = [uid for uid in emp_order if raw_records[uid]['punches']]

    for key, default in [
        ('holiday_dates', []),
        ('wfh_records',   {}),
        ('fixes',         {}),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    with st.sidebar:
        st.header("⚙️ Settings")
        target_weekly = st.number_input(
            "Full-Time Weekly Target (hrs, 6-day week)",
            min_value=1.0, value=51.0, step=0.5
        )
        daily_target = round(target_weekly / 6, 10)

        st.divider()
        st.subheader("🕑 Part-Time Settings")
        pt_daily_target = st.number_input(
            "Part-Time Daily Target (hrs)", min_value=0.5, value=4.0, step=0.5
        )
        part_time_list = st.multiselect(
            "Select Part-Time Employees", options=active_employees
        )

        st.divider()
        st.subheader("🏖️ Office Holidays")
        st.caption("8.30 hrs credited as paid holiday; not counted as a working day.")

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

    holiday_dates = st.session_state.holiday_dates
    wfh_records   = st.session_state.wfh_records

    st.header("🔧 Fix Missing Punches")
    any_missing = False
    for uid in active_employees:
        p_dict    = raw_records[uid]['punches']
        emp_fixes = st.session_state.fixes.get(uid, {})
        for day, p in sorted(p_dict.items()):
            if len(p) == 1:
                any_missing = True
                c1, c2, c3, c4 = st.columns([2, 1, 3, 2])
                c1.markdown(f"**{uid}**")
                c2.write(f"Day {day}")
                h = int(p[0].split(':')[0]) if ':' in p[0] else 0
                if h >= 12:
                    c3.warning(f"Out: {p[0]} (In missing)")
                    f_in = c4.text_input("Set In (HH:MM)", value="09:30", key=f"{uid}_{day}_in")
                    try:
                        datetime.strptime(f_in, "%H:%M")
                        emp_fixes[day] = {'in': f_in, 'out': p[0]}
                    except:
                        c4.error("Use HH:MM")
                else:
                    c3.warning(f"In: {p[0]} (Out missing)")
                    f_out = c4.text_input("Set Out (HH:MM)", value="18:00", key=f"{uid}_{day}_out")
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
        holiday_dates
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
        days_worked  = get_days_worked(uid, employees_dec, wfh_records, holiday_dates, year, month)
        leave_days   = get_leave_days(uid, raw_records, year, month, holiday_dates, wfh_records)
        hol_on_leave = get_holidays_on_leave(uid, raw_records, year, month, holiday_dates, wfh_records)

        week_dict = employees_dec[uid]
        for wk in sorted(week_dict.keys()):
            current_daily = pt_daily_target if uid in part_time_list else daily_target
            wk_target     = get_week_target(wk, year, month, current_daily)
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

    st.header("📥 Download Final Report")
    if st.button("Generate Excel Report", type="primary"):
        buf = generate_report(
            employees_dec, active_employees, raw_records, period_str,
            year, month, daily_target, part_time_list, pt_daily_target,
            holiday_dates, wfh_records
        )
        st.download_button(
            "⬇️ Download attendance_report.xlsx",
            buf, "attendance_report.xlsx"
        )

if __name__ == "__main__":
    main()
