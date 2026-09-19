# -*- coding: utf-8 -*-
"""Verify salary/attendance calculations in JUN.xlsx (or similar manual sheets).

This script reads each employee sheet, recomputes key totals from daily records,
and compares them against the summary values stored in the sheet.

Supported layouts (detected heuristically from header row):
    Type A: 4 cols [DAY, START, END, WORKING_HOURS] -> no explicit break
    Type B: 6 cols [DAY, START, END, GROSS, BREAK, NET]
    Type C: 7 cols part-time style with NET in column G/H
    Type D: 11 cols [START, REST, BACK, END, GROSS, STANDARD_REST, EXCEED_REST]
    Type E: 16 cols with multiple rest periods
"""
from __future__ import annotations

import argparse
import datetime
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.utils import get_column_letter


@dataclass
class DailyRecord:
    day: int
    start: Optional[datetime.time] = None
    end: Optional[datetime.time] = None
    gross_minutes: float = 0.0   # end - start
    break_minutes: float = 0.0   # explicit break
    net_minutes: float = 0.0     # gross - break


@dataclass
class SheetSummary:
    total_net_minutes: Optional[float] = None
    basic_minutes: Optional[float] = None
    ot_minutes: Optional[float] = None
    ph_work_minutes: Optional[float] = None
    ph_standard_minutes: Optional[float] = None
    ph_ot_minutes: Optional[float] = None
    oot_minutes: Optional[float] = None
    final_work_minutes: Optional[float] = None  # part-time


@dataclass
class VerificationResult:
    sheet_name: str
    layout_type: str
    is_full_time: Optional[bool] = None
    expected: SheetSummary = field(default_factory=SheetSummary)
    computed: SheetSummary = field(default_factory=SheetSummary)
    messages: List[str] = field(default_factory=list)


def time_to_minutes(t) -> Optional[float]:
    """Convert datetime.time or timedelta to minutes since midnight / duration."""
    if t is None:
        return None
    if isinstance(t, datetime.time):
        return t.hour * 60 + t.minute + t.second / 60.0
    if isinstance(t, datetime.timedelta):
        return t.total_seconds() / 60.0
    return None


def minutes_to_hours_str(mins: Optional[float]) -> str:
    if mins is None:
        return 'N/A'
    h = int(mins // 60)
    m = mins % 60
    return f'{h}h {m:05.2f}m ({mins/60:.2f}h)'


def td_to_minutes(td) -> Optional[float]:
    if isinstance(td, datetime.timedelta):
        return td.total_seconds() / 60.0
    return None


def classify_layout(header: List) -> str:
    """Classify sheet layout based on header row."""
    header_str = [str(h).upper().strip() if h else '' for h in header]
    n_cols = len([h for h in header_str if h])

    # Check keywords
    has_start = any('START' in h for h in header_str)
    has_end = any('END' in h for h in header_str)
    has_rest = any('REST' in h for h in header_str)
    has_back = any('BACK' in h for h in header_str)
    has_working = any('WORKING' in h for h in header_str)
    has_standard_rest = any('STANDARD REST' in h for h in header_str)
    has_exceed = any('EXCEED' in h for h in header_str)

    if n_cols >= 14 and has_start and has_rest and has_back and has_end and has_working:
        return 'E'  # 16-col multi-rest
    if has_start and has_rest and has_back and has_end and has_working:
        return 'D'  # 11-col single rest (or more with STANDARD REST / EXCEED)
    if has_start and has_end and has_working and n_cols <= 5:
        return 'A'  # 4-col simple
    if has_start and has_end and has_working and any('BREAK' in h or 'REST' in h for h in header_str):
        return 'B'  # 6-col start/end/break/net
    return 'C'  # fallback / part-time style


def parse_basic_hours_text(text: str) -> Optional[Tuple[float, float]]:
    """Parse text like 'JUN\\'26 - BASIC HOURS (26DAYS*7.5HRS) = 195HRS'.
    Returns (expected_days, basic_hours).
    """
    if not text:
        return None
    text = str(text).upper().replace(' ', '')
    m = re.search(r'(\d+(?:\.\d+)?)DAYS\*(\d+(?:\.\d+)?)HRS', text)
    if m:
        days = float(m.group(1))
        hours_per_day = float(m.group(2))
        return days, days * hours_per_day
    # Try just extracting the = XHRS part
    m2 = re.search(r'=(\d+(?:\.\d+)?)HRS', text)
    if m2:
        return None, float(m2.group(1))
    return None


def parse_public_holiday_text(text: str) -> Optional[Tuple[int, int, int]]:
    """Parse 'PUBLIC HOLIDAYS (01/06/2026)' -> (day, month, year)."""
    if not text:
        return None
    m = re.search(r'PUBLIC\s+HOLIDAYS\s*\((\d{1,2})/(\d{1,2})/(\d{4})\)', str(text), re.IGNORECASE)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    m2 = re.search(r'PUBLIC\s+HOLIDAYS\s*\((\d{1,2})/(\d{1,2})/(\d{2})\)', str(text), re.IGNORECASE)
    if m2:
        year = int(m2.group(3))
        year = 2000 + year if year < 50 else 1900 + year
        return int(m2.group(1)), int(m2.group(2)), year
    return None


def find_cell_by_keyword(ws, keyword: str, min_row: int = 1) -> Optional[Tuple[int, int, any]]:
    """Find first cell containing keyword (case-insensitive) at or after min_row."""
    for row_idx in range(min_row, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            if val is not None and keyword.upper() in str(val).upper():
                return row_idx, col_idx, val
    return None


def find_value_near_keyword(ws, keyword: str, offset_col: int = 1) -> Optional[float]:
    """Find keyword cell, then return numeric/timedelta value offset_col to the right."""
    found = find_cell_by_keyword(ws, keyword)
    if not found:
        return None
    row_idx, col_idx, _ = found
    target_col = col_idx + offset_col
    if target_col > ws.max_column:
        return None
    val = ws.cell(row=row_idx, column=target_col).value
    if isinstance(val, datetime.timedelta):
        return val.total_seconds() / 60.0
    if isinstance(val, (int, float)):
        return float(val)
    return None


def extract_expected_summary(ws, layout_type: str) -> Tuple[SheetSummary, bool, List[str]]:
    """Extract expected summary values from summary rows."""
    summary = SheetSummary()
    messages = []
    is_full_time = None

    # Detect full-time / part-time by presence of BASIC HOURS or OT
    has_basic = find_cell_by_keyword(ws, 'BASIC HOURS') is not None
    has_ot = find_cell_by_keyword(ws, 'OT') is not None
    has_total = find_cell_by_keyword(ws, 'TOTAL WORKING HOURS') is not None

    if has_basic or has_ot:
        is_full_time = True
    elif has_total and not has_basic:
        is_full_time = False  # likely part-time
    else:
        is_full_time = None

    # Total net work hours - try multiple sources
    # Source 1: month-specific "TOTAL WORKING HOURS (JUN'26)" or "TOTAL WORKING HOURS (MAY'26)"
    month_total_cell = None
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val and 'TOTAL WORKING HOURS' in str(val).upper() and '(' in str(val).upper():
                month_total_cell = (row_idx, col_idx)
                break
        if month_total_cell:
            break
    if month_total_cell:
        row_idx, col_idx = month_total_cell
        # Value is usually to the right (1 or 2 columns)
        for offset in [2, 3, 1]:
            if col_idx + offset > ws.max_column:
                continue
            val = ws.cell(row=row_idx, column=col_idx + offset).value
            if isinstance(val, datetime.timedelta):
                summary.total_net_minutes = val.total_seconds() / 60.0
                break

    # Source 2: generic "TOTAL WORKING HOURS" label
    if summary.total_net_minutes is None:
        val = find_value_near_keyword(ws, 'TOTAL WORKING HOURS', offset_col=1)
        if val is not None:
            summary.total_net_minutes = val

    # Source 3: first large numeric cell in summary rows (exclude OT/OOT rows)
    if summary.total_net_minutes is None:
        for row_idx in range(ws.max_row, ws.max_row - 7, -1):
            if row_idx < 1:
                continue
            # Skip rows containing OT/OOT keywords
            row_text = ' '.join(str(ws.cell(row=row_idx, column=c).value or '') for c in range(1, ws.max_column + 1))
            if ' OT ' in f' {row_text} '.upper() or 'OOT' in row_text.upper():
                continue
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if isinstance(val, datetime.timedelta) and val.total_seconds() > 0:
                    hours = val.total_seconds() / 3600
                    if hours > 20:
                        summary.total_net_minutes = val.total_seconds() / 60.0
                        break
            if summary.total_net_minutes is not None:
                break

    # Source 4: for full-time, total net = basic + OT
    if summary.total_net_minutes is None and summary.basic_minutes is not None and summary.ot_minutes is not None:
        summary.total_net_minutes = summary.basic_minutes + summary.ot_minutes

    # Basic hours
    basic_cell = find_cell_by_keyword(ws, 'BASIC HOURS')
    if basic_cell:
        text = str(basic_cell[2])
        parsed = parse_basic_hours_text(text)
        if parsed and parsed[1] is not None:
            summary.basic_minutes = parsed[1] * 60

    # OT
    summary.ot_minutes = find_value_near_keyword(ws, 'OT', offset_col=1)

    # Public holiday
    ph_cell = find_cell_by_keyword(ws, 'PUBLIC HOLIDAYS')
    if ph_cell:
        row_idx, col_idx, text = ph_cell
        # Work hours usually in column D or nearby
        for offset in [2, 3, 1, 4]:
            if col_idx + offset > ws.max_column:
                continue
            val = ws.cell(row=row_idx, column=col_idx + offset).value
            if isinstance(val, datetime.timedelta):
                summary.ph_work_minutes = val.total_seconds() / 60.0
                break
        # Standard hours usually a datetime.time like 7:30
        for offset in [3, 4, 2, 5]:
            if col_idx + offset > ws.max_column:
                continue
            val = ws.cell(row=row_idx, column=col_idx + offset).value
            if isinstance(val, datetime.time):
                summary.ph_standard_minutes = val.hour * 60 + val.minute
                break
            elif isinstance(val, datetime.timedelta):
                summary.ph_standard_minutes = val.total_seconds() / 60.0
                break

    # OOT
    summary.oot_minutes = find_value_near_keyword(ws, 'OOT', offset_col=1)

    # For full-time, total net = basic + OT is more reliable than heuristics
    if is_full_time and summary.basic_minutes is not None and summary.ot_minutes is not None:
        summary.total_net_minutes = summary.basic_minutes + summary.ot_minutes

    # Part-time final work hours - look for large timedelta in last few rows
    if is_full_time is False:
        for row_idx in range(ws.max_row, ws.max_row - 3, -1):
            if row_idx < 1:
                continue
            for col_idx in range(ws.max_column, 0, -1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if isinstance(val, datetime.timedelta) and val.total_seconds() > 0:
                    hours = val.total_seconds() / 3600
                    if hours > 20:
                        summary.final_work_minutes = val.total_seconds() / 60.0
                        break
            if summary.final_work_minutes is not None:
                break

    return summary, is_full_time, messages


def find_col_by_header(ws, keyword: str) -> Optional[int]:
    """Find column index (1-based) whose header contains keyword."""
    for col_idx in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=col_idx).value
        if h and keyword.upper() in str(h).upper():
            return col_idx
    return None


def find_candidate_net_column(ws, working_col: Optional[int]) -> Optional[int]:
    """Heuristically find a NET column to the right of working_col."""
    if not working_col:
        return None
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    # Look for explicit NET header
    for col_idx in range(ws.max_column, 0, -1):
        h = header[col_idx - 1]
        if h and 'NET' in str(h).upper():
            return col_idx
    # Otherwise, find rightmost column with timedelta values that are <= working hours
    best_col = None
    for col_idx in range(ws.max_column, working_col, -1):
        valid = 0
        smaller = 0
        for row_idx in range(2, min(ws.max_row, 33)):
            day = ws.cell(row=row_idx, column=1).value
            if not isinstance(day, int):
                continue
            work_val = td_to_minutes(ws.cell(row=row_idx, column=working_col).value)
            net_val = td_to_minutes(ws.cell(row=row_idx, column=col_idx).value)
            if net_val is None:
                continue
            valid += 1
            if work_val is not None and 0 < net_val <= work_val:
                smaller += 1
        if valid >= 3 and smaller >= valid * 0.5:
            best_col = col_idx
            break
    return best_col


def extract_daily_records(ws, layout_type: str) -> List[DailyRecord]:
    """Extract daily records based on layout type."""
    records = []
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

    # Find key columns by header name
    day_col = 1
    start_col = find_col_by_header(ws, 'START')
    end_col = find_col_by_header(ws, 'END')
    working_col = find_col_by_header(ws, 'WORKING')
    rest_col = find_col_by_header(ws, 'REST TIME') or find_col_by_header(ws, 'REST')  # daily rest start
    back_col = find_col_by_header(ws, 'BACK')
    standard_rest_col = find_col_by_header(ws, 'STANDARD REST') or find_col_by_header(ws, 'STANDARD TIME')
    exceed_rest_col = find_col_by_header(ws, 'EXCEED REST')

    # Determine net column (only for simple layouts; Type D/E uses gross - exceed_rest)
    has_rest_rules = standard_rest_col is not None or exceed_rest_col is not None
    net_col = None if has_rest_rules else find_candidate_net_column(ws, working_col)

    for row_idx in range(2, ws.max_row + 1):
        day = ws.cell(row=row_idx, column=day_col).value
        if not isinstance(day, int) or day < 1 or day > 31:
            continue

        rec = DailyRecord(day=day)
        if start_col:
            rec.start = ws.cell(row=row_idx, column=start_col).value
        if end_col:
            rec.end = ws.cell(row=row_idx, column=end_col).value

        # Gross working hours (in-company time)
        gross = None
        if working_col:
            gross = td_to_minutes(ws.cell(row=row_idx, column=working_col).value)
        if gross is None and start_col and end_col:
            start_mins = time_to_minutes(ws.cell(row=row_idx, column=start_col).value)
            end_mins = time_to_minutes(ws.cell(row=row_idx, column=end_col).value)
            if start_mins is not None and end_mins is not None:
                gross = end_mins - start_mins
        if gross is not None:
            rec.gross_minutes = gross

        # Net working hours
        net = None
        if net_col:
            net = td_to_minutes(ws.cell(row=row_idx, column=net_col).value)
        if net is None:
            net = gross
        if net is not None:
            rec.net_minutes = net

        # Break / rest calculation
        actual_rest = None
        # REST DURATION column (col E in LOW WEI SHENG)
        rest_duration_col = None
        for col_idx in range(1, ws.max_column + 1):
            h = header[col_idx - 1]
            if h is None:
                # Check if this column contains rest durations (timedelta between rest and back)
                pass
        # Use explicit rest duration if available
        if rest_col and back_col:
            rest_start = time_to_minutes(ws.cell(row=row_idx, column=rest_col).value)
            back_time = time_to_minutes(ws.cell(row=row_idx, column=back_col).value)
            if rest_start is not None and back_time is not None:
                actual_rest = back_time - rest_start
        # Also try col E pattern (rest duration column with no header)
        if actual_rest is None and rest_col and back_col:
            # Find column between REST and BACK
            for col_idx in range(min(rest_col, back_col) + 1, max(rest_col, back_col)):
                val = td_to_minutes(ws.cell(row=row_idx, column=col_idx).value)
                if val is not None and 0 < val < 480:
                    actual_rest = val
                    break

        standard_rest = None
        if standard_rest_col:
            standard_rest = td_to_minutes(ws.cell(row=row_idx, column=standard_rest_col).value)

        exceed_rest = None
        if exceed_rest_col:
            exceed_rest = td_to_minutes(ws.cell(row=row_idx, column=exceed_rest_col).value)
        elif actual_rest is not None and standard_rest is not None:
            exceed_rest = max(0.0, actual_rest - standard_rest)

        if exceed_rest is not None:
            rec.break_minutes = exceed_rest
            # For Type D/E, working hours is gross; net = gross - exceed_rest
            if gross is not None and net_col is None:
                rec.net_minutes = gross - exceed_rest

        records.append(rec)

    return records


def compute_summary(records: List[DailyRecord], expected: SheetSummary, is_full_time: Optional[bool], ph_date: Optional[Tuple[int, int, int]]) -> Tuple[SheetSummary, List[str]]:
    """Compute summary from daily records."""
    computed = SheetSummary()
    messages = []

    # Total net work (only days with positive net)
    total_net = sum(r.net_minutes for r in records if r.net_minutes > 0)
    computed.total_net_minutes = total_net

    # Public holiday work
    ph_work = 0.0
    if ph_date:
        ph_day = ph_date[0]
        for r in records:
            if r.day == ph_day and r.net_minutes > 0:
                ph_work = r.net_minutes
                break
    computed.ph_work_minutes = ph_work if ph_work > 0 else None
    computed.ph_standard_minutes = expected.ph_standard_minutes

    if is_full_time:
        # Basic hours
        computed.basic_minutes = expected.basic_minutes

        # OT = total net - basic
        if computed.basic_minutes is not None:
            computed.ot_minutes = total_net - computed.basic_minutes

        # Public holiday OT
        if computed.ph_work_minutes is not None and computed.ph_standard_minutes is not None:
            computed.ph_ot_minutes = computed.ph_work_minutes - computed.ph_standard_minutes

        # OOT = OT + PH OT
        if computed.ot_minutes is not None and computed.ph_ot_minutes is not None:
            computed.oot_minutes = computed.ot_minutes + computed.ph_ot_minutes
        elif computed.ot_minutes is not None:
            computed.oot_minutes = computed.ot_minutes

    else:
        # Part-time final = total net + public holiday work
        computed.final_work_minutes = total_net
        if computed.ph_work_minutes:
            computed.final_work_minutes += computed.ph_work_minutes

    return computed, messages


def compare_values(name: str, expected: Optional[float], computed: Optional[float], tolerance: float = 1.0) -> str:
    """Compare two minute values and return status string."""
    if expected is None and computed is None:
        return f'{name}: 两者皆无'
    if expected is None:
        return f'{name}: 文档无值，计算得 {minutes_to_hours_str(computed)}'
    if computed is None:
        return f'{name}: 文档为 {minutes_to_hours_str(expected)}，无法计算'
    diff = abs(expected - computed)
    status = '✅ 匹配' if diff <= tolerance else '❌ 不匹配'
    return (f'{name}: {status} | 文档={minutes_to_hours_str(expected)} | '
            f'计算={minutes_to_hours_str(computed)} | 差={diff:.1f}分钟')


def verify_sheet(ws, sheet_name: str) -> VerificationResult:
    """Verify one employee sheet."""
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    layout_type = classify_layout(header)

    result = VerificationResult(sheet_name=sheet_name, layout_type=layout_type)

    if layout_type in ('D', 'E'):
        result.messages.append(f'布局类型 {layout_type}（多休息段/复杂布局）暂不完全支持精确重算，以下仅做部分核对。')

    # Extract expected values
    expected, is_full_time, msgs = extract_expected_summary(ws, layout_type)
    result.expected = expected
    result.is_full_time = is_full_time
    result.messages.extend(msgs)

    # Extract daily records
    records = extract_daily_records(ws, layout_type)
    result.messages.append(f'解析到 {len(records)} 天打卡记录')

    # Public holiday date
    ph_cell = find_cell_by_keyword(ws, 'PUBLIC HOLIDAYS')
    ph_date = parse_public_holiday_text(ph_cell[2]) if ph_cell else None
    if ph_date:
        result.messages.append(f'检测到公共假期：{ph_date[0]:02d}/{ph_date[1]:02d}/{ph_date[2]}')

    # Compute
    computed, compute_msgs = compute_summary(records, expected, is_full_time, ph_date)
    result.computed = computed
    result.messages.extend(compute_msgs)

    return result


def verify_workbook(input_path: str) -> List[VerificationResult]:
    """Verify all sheets in workbook."""
    wb = openpyxl.load_workbook(input_path, data_only=True)
    results = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        result = verify_sheet(ws, sheet_name)
        results.append(result)
    return results


def print_report(results: List[VerificationResult]):
    """Print verification report."""
    print('=' * 80)
    print('Salary/Attendance Verification Report')
    print('=' * 80)

    for r in results:
        print(f'\n【{r.sheet_name}】 布局类型: {r.layout_type} | 类型: ', end='')
        if r.is_full_time is True:
            print('全职')
        elif r.is_full_time is False:
            print('兼职')
        else:
            print('未知')

        for msg in r.messages:
            print(f'  • {msg}')

        print()
        print('  ' + compare_values('总净工时', r.expected.total_net_minutes, r.computed.total_net_minutes))

        if r.is_full_time:
            print('  ' + compare_values('基本工时', r.expected.basic_minutes, r.computed.basic_minutes))
            print('  ' + compare_values('加班(OT)', r.expected.ot_minutes, r.computed.ot_minutes))
            print('  ' + compare_values('公共假期工时', r.expected.ph_work_minutes, r.computed.ph_work_minutes))
            print('  ' + compare_values('公共假期标准工时', r.expected.ph_standard_minutes, r.computed.ph_standard_minutes))
            print('  ' + compare_values('公共假期加班', r.expected.ph_ot_minutes, r.computed.ph_ot_minutes))
            print('  ' + compare_values('最终加班(OOT)', r.expected.oot_minutes, r.computed.oot_minutes))
        else:
            print('  ' + compare_values('最终工作时长', r.expected.final_work_minutes, r.computed.final_work_minutes))


def write_excel_report(results: List[VerificationResult], output_path: str):
    """Write verification results to an Excel file for easy review."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Verification Summary'

    headers = [
        'Employee', 'Layout', 'Type', 'Metric', 'Document Value', 'Computed Value',
        'Diff (min)', 'Status', 'Notes'
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='DDDDDD')

    green = PatternFill('solid', fgColor='C6EFCE')
    red = PatternFill('solid', fgColor='FFC7CE')

    for r in results:
        emp_type = 'Full-time' if r.is_full_time else ('Part-time' if r.is_full_time is False else 'Unknown')
        notes = ' | '.join(r.messages)

        metrics = [
            ('Total Net Hours', r.expected.total_net_minutes, r.computed.total_net_minutes),
        ]
        if r.is_full_time:
            metrics.extend([
                ('Basic Hours', r.expected.basic_minutes, r.computed.basic_minutes),
                ('OT', r.expected.ot_minutes, r.computed.ot_minutes),
                ('PH Work', r.expected.ph_work_minutes, r.computed.ph_work_minutes),
                ('PH Standard', r.expected.ph_standard_minutes, r.computed.ph_standard_minutes),
                ('PH OT', r.expected.ph_ot_minutes, r.computed.ph_ot_minutes),
                ('OOT', r.expected.oot_minutes, r.computed.oot_minutes),
            ])
        else:
            metrics.append(('Final Work Hours', r.expected.final_work_minutes, r.computed.final_work_minutes))

        for metric_name, exp, comp in metrics:
            if exp is None and comp is None:
                continue
            diff = None if exp is None or comp is None else abs(exp - comp)
            status = 'Match' if diff is not None and diff <= 1.0 else ('Mismatch' if diff is not None else 'N/A')
            ws.append([
                r.sheet_name,
                r.layout_type,
                emp_type,
                metric_name,
                minutes_to_hours_str(exp),
                minutes_to_hours_str(comp),
                f'{diff:.1f}' if diff is not None else 'N/A',
                status,
                notes
            ])
            row_idx = ws.max_row
            fill = green if status == 'Match' else (red if status == 'Mismatch' else None)
            if fill:
                for cell in ws[row_idx]:
                    cell.fill = fill

    ws.freeze_panes = 'A2'
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[column].width = min(max_length + 2, 50)

    wb.save(output_path)


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description='Verify salary/attendance calculations in manual Excel')
    parser.add_argument('input', help='Input .xlsx file (e.g., JUN.xlsx)')
    parser.add_argument('-o', '--output', help='Output Excel report path')
    args = parser.parse_args(argv)

    if not Path(args.input).exists():
        print(f'Error: file not found: {args.input}', file=sys.stderr)
        return 1

    results = verify_workbook(args.input)
    print_report(results)

    if args.output:
        write_excel_report(results, args.output)
        print(f'\nExcel report saved to: {args.output}')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
