import io
from datetime import datetime, date
from collections import defaultdict

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
WEEKLY_TARGET = 51.0   # hours per week

# ─────────────────────────────────────────────────────────────────────────────
# PARSING LOGIC
# ─────────────────────────────────────────────────────────────────────────────

def parse_punches(cell_value):
    if not cell_value:
        return []
    return [t.strip() for t in str(cell_value).split('\n') if t.strip()]

def minutes_to_hhmm(total_minutes):
    h = int(total_minutes // 60)
    m = int(total_minutes % 60)
    return f"{h}.{m:02d}"

def decimal_to_hhmm(decimal_hours):
    total_minutes = round(decimal_hours * 60)
    return minutes_to_hhmm(total_minutes)

def compute_hours_from_pair(t_in_str, t_out_str):
    try:
        t_in  = datetime.strptime(t_in_str,  "%H:%M")
        t_out = datetime.strptime(t_out_str, "%H:%M")
        diff_min = int((t_out - t_in).total_seconds() / 60)
        if diff_min <= 0:
            return 0.0, "0.00"
        return round(diff_min / 60, 4), minutes_to_hhmm(diff_min)
    except ValueError:
        return 0.0, "0.00"

def get_week_number(day, year, month):
    week_count = 1
    first_weekday = date(year, month, 1).weekday()  # 0=Mon, 6=Sun

    for d in range(1, day + 1):
        if date(year, month, d).weekday() == 0 and d != 1:
            # Only bump the week if this Monday is NOT immediately
            # following a Sunday start (i.e., day 2 when day 1 is Sunday)
            if not (d == 2 and first_weekday == 6):
                week_count += 1
    return week_count

def get_week_target(relative_wk, year, month, daily_target):
    """Target hours based on Mon-Sat days of the relative calendar week."""
    working_days = 0
    for d_int in range(1, 32):
        try:
            if get_week_number(d_int, year, month) == relative_wk:
                d_obj = date(year, month, d_int)
                if d_obj.weekday() != 6:  # Not Sunday
                    working_days += 1
        except ValueError:
            break
    return round(working_days * daily_target, 2)

def parse_logs_sheet(ws):
    all_rows = list(ws.iter_rows(values_only=True))
    period_str = ""
    year, month = datetime.now().year, datetime.now().month
    for row in all_rows[:5]:
        for cell in row:
            if cell and isinstance(cell, str) and '~' in cell:
                period_str = cell.strip()
                try:
                    start_part = period_str.split('~')[0].strip()
                    dt = datetime.strptime(start_part, "%Y/%m/%d")
                    year, month = dt.year, dt.month
                except Exception: pass

    raw_records = {}
    emp_order   = []

    i = 0
    while i < len(all_rows):
        row = all_rows[i]
        if row and row[0] == 'No :':
            emp_no   = str(row[2]).strip()  if row[2]  else 'Unknown'
            emp_name = str(row[10]).strip() if row[10] else 'Unnamed'
            unique_id = f"{emp_name.title()} (ID: {emp_no})"

            days_row      = all_rows[i - 1] if i > 0 else []
            punch_row_idx = i + 1

            if punch_row_idx < len(all_rows):
                punch_row   = all_rows[punch_row_idx]
                day_punches = {}

                for col, day_num in enumerate(days_row):
                    if not isinstance(day_num, int):
                        continue
                    if col >= len(punch_row):
                        continue
                    try:
                        if date(year, month, day_num).weekday() == 6:
                            continue
                    except Exception:
                        continue
                    
                    punches = parse_punches(punch_row[col])
                    if punches:
                        day_punches[day_num] = punches

                if unique_id not in raw_records:
                    raw_records[unique_id] = {'name': emp_name, 'id': emp_no, 'punches': {}}
                    emp_order.append(unique_id)
                raw_records[unique_id]['punches'].update(day_punches)
        i += 1
    return raw_records, emp_order, period_str, year, month

# ─────────────────────────────────────────────────────────────────────────────
# HOUR CALCULATIONS
# ─────────────────────────────────────────────────────────────────────────────

def sum_week_hours(day_dict):
    return sum(day_dict.values())

def calculate_hour_metrics(week_hours_decimal, target_weekly):
    excess   = max(0.0, week_hours_decimal - target_weekly)
    shortage = max(0.0, target_weekly - week_hours_decimal)
    return {
        'hours_worked_display': decimal_to_hhmm(week_hours_decimal),
        'target_display':       decimal_to_hhmm(target_weekly),
        'excess_hours':         excess,
        'excess_display':       decimal_to_hhmm(excess),
        'shortage_hours':       shortage,
        'shortage_display':     decimal_to_hhmm(shortage),
    }

# ─────────────────────────────────────────────────────────────────────────────
# STYLING & EXCEL WRITING
# ─────────────────────────────────────────────────────────────────────────────

C_HEADER_BG, C_HEADER_FG = "1F4E79", "FFFFFF"
C_SUBHDR_BG, C_EXCESS_BG = "2E75B6", "E2EFDA"
C_SHORT_BG, C_NEUTRAL_BG = "FCE4D6", "DEEAF1"
C_ALT_ROW, C_BORDER      = "F2F7FB", "BDD7EE"

def make_border():
    s = Side(style='thin', color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def make_font(bold=False, color="000000", size=10):
    return Font(name='Arial', bold=bold, color=color, size=size)

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_align(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def write_summary_sheet(wb, employees_dec, emp_order, raw_records, period_str, target_weekly, year, month, daily_target):
    ws = wb.create_sheet("Weekly Summary")
    ws.merge_cells("A1:H1")
    ws["A1"] = f"Weekly Attendance Breaks | {period_str}"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = make_font(True, C_HEADER_FG, 14), make_fill(C_HEADER_BG), make_align()

    headers = ["ID", "Employee Name", "Week", "Hours Worked", "Target Hours", "Excess", "Shortage", "Status"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()

    row = 4
    for uid in emp_order:
        week_dict = employees_dec.get(uid, {})
        for wk in sorted(week_dict.keys()):
            wk_target = get_week_target(wk, year, month, daily_target)
            met = calculate_hour_metrics(sum_week_hours(week_dict[wk]), wk_target)
            status = "EXCESS" if met['excess_hours'] > 0 else "SHORTAGE" if met['shortage_hours'] > 0 else "ON TARGET"
            row_fill = make_fill(C_EXCESS_BG if met['excess_hours'] > 0 else C_SHORT_BG if met['shortage_hours'] > 0 else C_ALT_ROW)
            vals = [raw_records[uid]['id'], raw_records[uid]['name'].title(), f"Week {wk}",
                    met['hours_worked_display'], met['target_display'], met['excess_display'], met['shortage_display'], status]
            for col, v in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill, c.border, c.alignment, c.font = row_fill, make_border(), make_align(), make_font()
            row += 1
    ws.column_dimensions['B'].width = 25

def write_consolidated_sheet(wb, employees_dec, emp_order, raw_records, period_str, target_weekly, year, month, daily_target):
    ws = wb.create_sheet("Consolidated Report")
    ws.merge_cells("A1:F1")
    ws["A1"] = f"Total Monthly Consolidation | {period_str}"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = make_font(True, C_HEADER_FG, 14), make_fill(C_HEADER_BG), make_align()

    headers = ["ID", "Employee Name", "Total Hours", "Total Target", "Total Excess", "Total Shortage"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()

    row = 4
    for uid in emp_order:
        if uid in employees_dec:
            week_dict = employees_dec[uid]
            total_hrs      = sum(sum_week_hours(d) for d in week_dict.values())
            total_target   = 0.0
            total_excess   = 0.0
            total_shortage = 0.0
            for wk in week_dict:
                wk_target = get_week_target(wk, year, month, daily_target)
                total_target  += wk_target
                met = calculate_hour_metrics(sum_week_hours(week_dict[wk]), wk_target)
                total_excess   += met['excess_hours']
                total_shortage += met['shortage_hours']

            vals = [raw_records[uid]['id'], raw_records[uid]['name'].title(), decimal_to_hhmm(total_hrs),
                    decimal_to_hhmm(total_target), decimal_to_hhmm(total_excess), decimal_to_hhmm(total_shortage)]
            for col, v in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill, c.border, c.alignment, c.font = make_fill(C_ALT_ROW), make_border(), make_align(), make_font()
            row += 1
    ws.column_dimensions['B'].width = 25
    for col in ['C', 'D', 'E', 'F']: ws.column_dimensions[col].width = 15

def write_individual_sheet(wb, uid, week_dict, period_str, target_weekly, year, month, daily_target):
    ws_name = uid[:30].replace(":", "").replace("/", "").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
    ws = wb.create_sheet(ws_name)

    ws.merge_cells("A1:C1")
    ws["A1"] = f"Attendance Details: {uid}"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = make_font(True, C_HEADER_FG, 12), make_fill(C_HEADER_BG), make_align()

    row = 3
    for wk in sorted(week_dict.keys()):
        ws.merge_cells(f"A{row}:C{row}")
        c_wk = ws.cell(row=row, column=1, value=f"WEEK {wk}")
        c_wk.font, c_wk.fill, c_wk.alignment = make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align()
        row += 1

        for col, h in enumerate(["Date", "Day", "Hours Worked"], 1):
            c = ws.cell(row=row, column=col, value=h)
            c.font, c.fill, c.alignment, c.border = make_font(True, C_HEADER_FG), make_fill(C_SUBHDR_BG), make_align(), make_border()
        row += 1

        for day, hrs in sorted(week_dict[wk].items()):
            try:
                dt_obj = date(year, month, day)
                d_str = dt_obj.strftime("%d-%b-%Y")
                d_name = dt_obj.strftime("%A")
            except:
                d_str, d_name = f"Day {day}", ""

            for col, v in enumerate([d_str, d_name, decimal_to_hhmm(hrs)], 1):
                c = ws.cell(row=row, column=col, value=v)
                c.fill, c.border, c.alignment, c.font = make_fill(C_ALT_ROW), make_border(), make_align(), make_font()
            row += 1

        wk_target = get_week_target(wk, year, month, daily_target)
        met = calculate_hour_metrics(sum_week_hours(week_dict[wk]), wk_target)
        summary_fill = make_fill(C_NEUTRAL_BG)

        for label, val in [
            ("Total Worked (Week)", met['hours_worked_display']),
            ("Target (Week)", met['target_display']),
            ("Excess Hours", met['excess_display']),
            ("Shortage Hours", met['shortage_display'])
        ]:
            ws.merge_cells(f"A{row}:B{row}")
            c_lbl = ws.cell(row=row, column=1, value=label)
            c_val = ws.cell(row=row, column=3, value=val)
            for c in [c_lbl, c_val]:
                c.font, c.fill, c.border, c.alignment = make_font(True), summary_fill, make_border(), make_align()
            row += 1
        row += 1 

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 18

def generate_report(employees_dec, emp_order, raw_records, period_str, year, month, target_weekly, daily_target):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_consolidated_sheet(wb, employees_dec, emp_order, raw_records, period_str, target_weekly, year, month, daily_target)
    write_summary_sheet(wb, employees_dec, emp_order, raw_records, period_str, target_weekly, year, month, daily_target)
    for uid in emp_order:
        if uid in employees_dec:
            write_individual_sheet(wb, uid, employees_dec[uid], period_str, target_weekly, year, month, daily_target)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Attendance Processor", layout="wide")
    st.title("🕐 Attendance Processor (Full Individual Reports)")

    with st.sidebar:
        st.header("Settings")
        target_weekly = st.number_input("Weekly Target Hours (6-day week)", min_value=1.0, value=51.0, step=0.5)
        daily_target  = round(target_weekly / 6, 10)

    uploaded = st.file_uploader("📂 Upload Attendance XLSX", type=["xlsx", "xls"])
    if not uploaded: return

    wb_in = openpyxl.load_workbook(uploaded, read_only=True)
    log_sheet = next((wb_in[n] for n in wb_in.sheetnames if n.lower() == 'logs'), None)
    if not log_sheet:
        st.error("No 'Logs' sheet found."); return

    raw_records, emp_order, period_str, year, month = parse_logs_sheet(log_sheet)

    if 'fixes' not in st.session_state: st.session_state.fixes = {}
    st.header("🔧 Fix Missing Punches")
    for uid in emp_order:
        p_dict = raw_records[uid]['punches']
        emp_fixes = st.session_state.fixes.get(uid, {})
        for day, p in sorted(p_dict.items()):
            if len(p) == 1:
                col1, col2, col3, col4 = st.columns([2, 1, 3, 2])
                col1.markdown(f"**{uid}**")
                col2.write(f"Day {day}")
                h = int(p[0].split(':')[0]) if ':' in p[0] else 0
                if h >= 12:
                    col3.warning(f"Out: {p[0]} (In Missing)")
                    f_in = col4.text_input("Set In (HH:MM)", value="09:30", key=f"{uid}_{day}_in")
                    try:
                        datetime.strptime(f_in, "%H:%M")
                        emp_fixes[day] = {'in': f_in, 'out': p[0]}
                    except ValueError:
                        col4.error("Use HH:MM format")
                else:
                    col3.warning(f"In: {p[0]} (Out Missing)")
                    f_out = col4.text_input("Set Out (HH:MM)", value="18:00", key=f"{uid}_{day}_out")
                    try:
                        datetime.strptime(f_out, "%H:%M")
                        emp_fixes[day] = {'in': p[0], 'out': f_out}
                    except ValueError:
                        col4.error("Use HH:MM format")
        if emp_fixes: st.session_state.fixes[uid] = emp_fixes

    employees_dec = {}
    for uid in emp_order:
        p_dict, f_dict, week_data = raw_records[uid]['punches'], st.session_state.fixes.get(uid, {}), defaultdict(dict)
        for day, p in p_dict.items():
            if day in f_dict: dec, _ = compute_hours_from_pair(f_dict[day]['in'], f_dict[day]['out'])
            elif len(p) >= 2: dec = sum(compute_hours_from_pair(p[i], p[i+1])[0] for i in range(0, len(p)-1, 2))
            elif len(p) == 1:
                h = int(p[0].split(':')[0]) if ':' in p[0] else 0
                dec, _ = compute_hours_from_pair("09:30", p[0]) if h >= 12 else compute_hours_from_pair(p[0], "18:00")
            else: dec = 0.0
            
            if dec > 0: 
                wk_id = get_week_number(day, year, month)
                week_data[wk_id][day] = dec
                
        if week_data: employees_dec[uid] = dict(week_data)

    st.header("📥 Download Final Report")
    st.info("The generated file will contain a Consolidated Summary sheet, a Weekly Summary sheet, and one tab for every Employee.")
    if st.button("Generate Excel Report", type="primary"):
        buf = generate_report(employees_dec, emp_order, raw_records, period_str, year, month, target_weekly, daily_target)
        st.download_button("⬇️ Download attendance_report.xlsx", buf, "attendance_report.xlsx")

if __name__ == "__main__": main()
