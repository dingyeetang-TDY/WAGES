#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parse attendance punch-clock data from an .xls/.xlsx Logs sheet to JSON.

Reads the 'Logs' sheet, extracts per-employee daily punch times, applies
odd/even segment rules, and writes a JSON file for downstream report generation.

Usage:
    python parse_logs.py input.xlsx
    python parse_logs.py input.xls -o parsed.json --month 7
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd


def parse_time(s: str) -> Optional[str]:
    s = str(s).strip()
    if s in ('nan', '', 'None'):
        return None
    m = re.match(r'(\d{1,2}):(\d{2})', s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return None


def time_to_minutes(t: str) -> int:
    h, m = map(int, t.split(':'))
    return h * 60 + m


def fmt_hours(mins: float) -> str:
    if mins <= 0:
        return '0.00'
    return f'{mins / 60:.2f}'


def detect_dates(df: pd.DataFrame, month: int = 7) -> List[str]:
    dates = []
    if df.shape[0] > 3:
        for col in range(1, df.shape[1]):
            val = df.iloc[3, col]
            if not pd.isna(val) and isinstance(val, (int, float)):
                dates.append(f'{month:02d}/{int(val):02d}')
    if not dates:
        dates = [f'{month:02d}/{d:02d}' for d in range(1, 19)]
    return dates


def parse_employees(df: pd.DataFrame, dates: List[str]) -> list:
    employees_logs = []
    for r in range(df.shape[0]):
        v0 = str(df.iloc[r, 0]).strip() if not pd.isna(df.iloc[r, 0]) else ''
        if v0 == 'No :':
            no = str(df.iloc[r, 2]).strip() if not pd.isna(df.iloc[r, 2]) else ''
            name = str(df.iloc[r, 10]).strip() if df.shape[1] > 10 and not pd.isna(df.iloc[r, 10]) else ''
            dept = str(df.iloc[r, 19]).strip() if df.shape[1] > 19 and not pd.isna(df.iloc[r, 19]) else ''
            employees_logs.append({'no': no, 'name': name, 'dept': dept, 'data_row': r + 1})

    all_employees = []
    for emp in employees_logs:
        r = emp['data_row']
        if r >= df.shape[0]:
            continue
        daily_records = {}
        for day_idx, date_str in enumerate(dates):
            col = day_idx + 1
            if col >= df.shape[1]:
                daily_records[date_str] = {'times': [], 'status': 'No Data'}
                continue
            val = df.iloc[r, col]
            if pd.isna(val):
                daily_records[date_str] = {'times': [], 'status': 'No Data'}
                continue
            val_str = str(val).strip()
            if val_str in ('nan', '', 'None'):
                daily_records[date_str] = {'times': [], 'status': 'No Data'}
                continue
            time_strs = [t.strip() for t in val_str.split('\n') if t.strip()]
            times = sorted([t for t in (parse_time(t) for t in time_strs) if t])
            daily_records[date_str] = {'times': times, 'status': 'Present'} if times else {'times': [], 'status': 'No Data'}

        all_employees.append({
            'no': emp['no'], 'name': emp['name'], 'dept': emp['dept'],
            'records': daily_records,
        })
    return all_employees


def apply_rules(employees: list) -> None:
    for emp in employees:
        emp['odd_annotations'] = {}
        emp['segments'] = {}
        emp['daily_work_mins'] = {}

        for date_str, rec in emp['records'].items():
            times = rec['times']
            n = len(times)

            if n == 0:
                emp['daily_work_mins'][date_str] = 0
                emp['segments'][date_str] = []
                continue

            if n % 2 != 0:
                emp['odd_annotations'][date_str] = n

            segments = []
            work_mins = 0

            if n == 2:
                m1, m2 = time_to_minutes(times[0]), time_to_minutes(times[1])
                mins = m2 - m1
                segments.append((times[0], times[1], '上午班', mins))
                work_mins += mins
            elif n == 4:
                m = [time_to_minutes(t) for t in times]
                segments.append((times[0], times[1], '上午班', m[1] - m[0]))
                segments.append((times[1], times[2], '中场休息', m[2] - m[1]))
                segments.append((times[2], times[3], '下午班', m[3] - m[2]))
                work_mins = (m[1] - m[0]) + (m[3] - m[2])
            elif n == 6:
                m = [time_to_minutes(t) for t in times]
                segments.append((times[0], times[1], '上午班', m[1] - m[0]))
                segments.append((times[1], times[2], '中场休息', m[2] - m[1]))
                segments.append((times[2], times[3], '下午班', m[3] - m[2]))
                segments.append((times[3], times[4], '中场休息2', m[4] - m[3]))
                segments.append((times[4], times[5], '下午班2', m[5] - m[4]))
                work_mins = (m[1] - m[0]) + (m[3] - m[2]) + (m[5] - m[4])
            elif n > 6 and n % 2 == 0:
                m = [time_to_minutes(t) for t in times]
                for i in range(0, n, 2):
                    seg_mins = m[i + 1] - m[i]
                    if i == 0:
                        seg_type = '上午班'
                    elif i == n - 2:
                        seg_type = '下午班2'
                    elif i == 2:
                        seg_type = '下午班'
                    else:
                        seg_type = f'中场休息{i // 2}'
                    segments.append((times[i], times[i + 1], seg_type, seg_mins))
                    if seg_type.startswith('上午') or seg_type.startswith('下午'):
                        work_mins += seg_mins
            else:
                m = [time_to_minutes(t) for t in times]
                for i in range(n - 1):
                    seg_mins = m[i + 1] - m[i]
                    seg_type = {0: '上午班', 1: '中场休息', 2: '下午班', 3: '中场休息2', 4: '下午班2'}.get(i, '未知')
                    segments.append((times[i], times[i + 1], seg_type, seg_mins))
                    if seg_type.startswith('上午') or seg_type.startswith('下午'):
                        work_mins += seg_mins

            emp['segments'][date_str] = segments
            emp['daily_work_mins'][date_str] = work_mins


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description='Parse attendance Logs sheet to JSON with work-hour rules applied.'
    )
    parser.add_argument('input', help='Input .xls or .xlsx file path')
    parser.add_argument('-o', '--output', default=None, help='Output .json file path')
    parser.add_argument('--sheet', default='Logs', help='Sheet name (default: Logs)')
    parser.add_argument('--month', type=int, default=7, help='Month for date labels (default: 7)')
    args = parser.parse_args(argv)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'Error: file not found: {input_path}', file=sys.stderr)
        return 1

    output_path = args.output or str(input_path.with_suffix('.json'))

    engine = 'xlrd' if str(input_path).lower().endswith('.xls') else None
    df = pd.read_excel(str(input_path), sheet_name=args.sheet, header=None, engine=engine) if engine \
        else pd.read_excel(str(input_path), sheet_name=args.sheet, header=None)

    dates = detect_dates(df, month=args.month)
    employees = parse_employees(df, dates)
    apply_rules(employees)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)

    total_odd = sum(len(e['odd_annotations']) for e in employees)
    total_work = sum(sum(e['daily_work_mins'].values()) for e in employees)
    print(f'JSON saved: {output_path}')
    print(f'Employees: {len(employees)}')
    print(f'Date columns: {len(dates)}')
    print(f'Odd-punch anomaly days: {total_odd}')
    print(f'Total work hours: {fmt_hours(total_work)}h')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
