#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Apply visual styling to an attendance .xlsx workbook.

Reads a plain .xlsx and writes a styled copy with:
  - Title rows: dark background + white bold text
  - Header rows: dark background + white bold text
  - Data rows: zebra striping (alternating white / light blue)
  - Absence cells: light red background + dark red italic text
  - Schedule cells: light yellow background + brown italic text
  - Log day-number rows: light blue background
  - Employee name rows: light gray background

NOTE: This script expects a specific workbook layout (Summary, Logs, and
date-grouped sheets like '21.29.31', '1.2.4', etc.). Adjust the sheet
detection logic in the `style_sheet()` function if your layout differs.

Usage:
    python beautify_xlsx.py input.xlsx
    python beautify_xlsx.py input.xlsx -o styled.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def style_summary_sheet(ws):
    """Apply styling to the Summary sheet."""
    mc = ws.max_column
    mr = ws.max_row
    TITLE_FILL = PatternFill('solid', fgColor='1F2937')
    TITLE_FONT = Font(name='Arial', bold=True, size=14, color='FFFFFF')
    SUBTITLE_FONT = Font(name='Arial', size=11, color='4B5563')
    HEADER_FILL = PatternFill('solid', fgColor='2C3E50')
    HEADER_FONT = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    DATA_FONT = Font(name='Arial', size=10, color='1F2937')
    ZEBRA1 = PatternFill('solid', fgColor='FFFFFF')
    ZEBRA2 = PatternFill('solid', fgColor='F7F9FC')
    THIN = Border(
        left=Side('thin', 'D9DEE7'), right=Side('thin', 'D9DEE7'),
        top=Side('thin', 'D9DEE7'), bottom=Side('thin', 'D9DEE7'))
    CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for c in range(1, mc + 1):
        ws.cell(1, c).font = TITLE_FONT
        ws.cell(1, c).fill = TITLE_FILL
        ws.cell(1, c).alignment = CENTER
        ws.cell(2, c).font = SUBTITLE_FONT
    for r in [3, 4]:
        for c in range(1, mc + 1):
            ws.cell(r, c).font = HEADER_FONT
            ws.cell(r, c).fill = HEADER_FILL
            ws.cell(r, c).alignment = CENTER
            ws.cell(r, c).border = THIN
    for r in range(5, mr + 1):
        for c in range(1, mc + 1):
            ws.cell(r, c).font = DATA_FONT
            ws.cell(r, c).fill = ZEBRA1 if (r - 5) % 2 == 0 else ZEBRA2
            ws.cell(r, c).border = THIN
            ws.cell(r, c).alignment = CENTER
    ws.row_dimensions[1].height = 30


def style_logs_sheet(ws):
    """Apply styling to the Logs sheet."""
    mc = ws.max_column
    mr = ws.max_row
    TITLE_FILL = PatternFill('solid', fgColor='1F2937')
    TITLE_FONT = Font(name='Arial', bold=True, size=14, color='FFFFFF')
    INFO_FILL = PatternFill('solid', fgColor='F0F4F8')
    INFO_FONT = Font(name='Arial', size=11, bold=True, color='1F2937')
    LOG_DAY_FILL = PatternFill('solid', fgColor='DBEAFE')
    LOG_DAY_FONT = Font(name='Arial', size=10, bold=True, color='1E40AF')
    LOG_NAME_FILL = PatternFill('solid', fgColor='F1F5F9')
    LOG_NAME_FONT = Font(name='Arial', size=10, bold=True, color='334155')
    LOG_DATA_FONT = Font(name='Arial', size=10, color='374151')
    THIN = Border(
        left=Side('thin', 'D9DEE7'), right=Side('thin', 'D9DEE7'),
        top=Side('thin', 'D9DEE7'), bottom=Side('thin', 'D9DEE7'))
    CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

    for c in range(1, mc + 1):
        ws.cell(1, c).font = TITLE_FONT
        ws.cell(1, c).fill = TITLE_FILL
        ws.cell(1, c).alignment = CENTER
        ws.cell(3, c).font = INFO_FONT
        ws.cell(3, c).fill = INFO_FILL

    for r in range(4, mr + 1):
        v0 = str(ws.cell(r, 1).value or '').strip()
        v1 = str(ws.cell(r, 2).value or '').strip()
        if v0.replace('.', '').isdigit() and v1.replace('.', '').isdigit():
            for c in range(1, mc + 1):
                cell = ws.cell(r, c)
                cell.font = LOG_DAY_FONT
                cell.fill = LOG_DAY_FILL
                cell.alignment = CENTER
                cell.border = THIN
        elif v0 == 'No :':
            for c in range(1, mc + 1):
                cell = ws.cell(r, c)
                cell.font = LOG_NAME_FONT
                cell.fill = LOG_NAME_FILL
                cell.alignment = LEFT
                cell.border = THIN
        else:
            for c in range(1, mc + 1):
                cell = ws.cell(r, c)
                cell.font = LOG_DATA_FONT
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                cell.border = THIN
    for c in range(1, mc + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16
    ws.row_dimensions[1].height = 30


def style_date_grouped_sheet(ws):
    """Apply styling to date-grouped sheets (e.g. '21.29.31', '1.2.4')."""
    mc = ws.max_column
    mr = ws.max_row

    TITLE_FILL = PatternFill('solid', fgColor='1F2937')
    TITLE_FONT = Font(name='Arial', bold=True, size=14, color='FFFFFF')
    INFO_FILL = PatternFill('solid', fgColor='F0F4F8')
    INFO_FONT = Font(name='Arial', size=11, bold=True, color='1F2937')
    HEADER_FILL = PatternFill('solid', fgColor='2C3E50')
    HEADER_FONT = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    DATA_FONT = Font(name='Arial', size=10, color='1F2937')
    DATA_FONT_BOLD = Font(name='Arial', size=10, bold=True, color='1F2937')
    ZEBRA1 = PatternFill('solid', fgColor='FFFFFF')
    ZEBRA2 = PatternFill('solid', fgColor='F7F9FC')
    ABSENT_FILL = PatternFill('solid', fgColor='FEE2E2')
    ABSENT_FONT = Font(name='Arial', size=10, color='991B1B', italic=True)
    THIN = Border(
        left=Side('thin', 'D9DEE7'), right=Side('thin', 'D9DEE7'),
        top=Side('thin', 'D9DEE7'), bottom=Side('thin', 'D9DEE7'))
    CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    SCHEDULE_FILL = PatternFill('solid', fgColor='FEF3C7')
    SCHEDULE_FONT = Font(name='Arial', size=10, color='92400E', italic=True)
    SUBHEADER_FILL = PatternFill('solid', fgColor='3D5A80')
    SUBHEADER_FONT = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    STATS_FILL = PatternFill('solid', fgColor='E8F0FE')
    TABLE_LABEL_FILL = PatternFill('solid', fgColor='E0E7EF')
    TABLE_LABEL_FONT = Font(name='Arial', bold=True, size=11, color='1E3A5F')

    for offset in [0, 15, 30]:
        end_c = min(offset + 14, mc)
        if offset + 1 > mc:
            continue
        for c in range(offset + 1, end_c + 1):
            ws.cell(1, c).font = TITLE_FONT
            ws.cell(1, c).fill = TITLE_FILL
            ws.cell(1, c).alignment = CENTER
        for r in [2, 3]:
            for c in range(offset + 1, end_c + 1):
                ws.cell(r, c).font = INFO_FONT
                ws.cell(r, c).fill = INFO_FILL
                ws.cell(r, c).alignment = LEFT
        for r in [4, 5, 10]:
            for c in range(offset + 1, end_c + 1):
                ws.cell(r, c).font = HEADER_FONT
                ws.cell(r, c).fill = HEADER_FILL
                ws.cell(r, c).alignment = CENTER
                ws.cell(r, c).border = THIN
        for c in range(offset + 1, end_c + 1):
            ws.cell(6, c).font = DATA_FONT_BOLD
            ws.cell(6, c).fill = STATS_FILL
            ws.cell(6, c).alignment = CENTER
            ws.cell(6, c).border = THIN
            ws.cell(7, c).font = SCHEDULE_FONT
            ws.cell(7, c).fill = SCHEDULE_FILL
            ws.cell(7, c).alignment = LEFT
            ws.cell(7, c).border = THIN
            ws.cell(8, c).font = TABLE_LABEL_FONT
            ws.cell(8, c).fill = TABLE_LABEL_FILL
            ws.cell(8, c).alignment = CENTER
            ws.cell(8, c).border = THIN
            ws.cell(9, c).font = SUBHEADER_FONT
            ws.cell(9, c).fill = SUBHEADER_FILL
            ws.cell(9, c).alignment = CENTER
            ws.cell(9, c).border = THIN
        for r in range(11, 29):
            if r > mr:
                break
            for c in range(offset + 1, end_c + 1):
                cell = ws.cell(r, c)
                v = str(cell.value or '').strip()
                if v == 'Absence':
                    cell.font = ABSENT_FONT
                    cell.fill = ABSENT_FILL
                else:
                    cell.font = DATA_FONT
                    cell.fill = ZEBRA1 if (r - 11) % 2 == 0 else ZEBRA2
                cell.alignment = CENTER
                cell.border = THIN
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[8].height = 22
    for r in range(11, 29):
        ws.row_dimensions[r].height = 20


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description='Apply visual styling to an attendance .xlsx workbook.'
    )
    parser.add_argument('input', help='Input .xlsx file path')
    parser.add_argument('-o', '--output', default=None, help='Output .xlsx file path')
    args = parser.parse_args(argv)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'Error: file not found: {input_path}', file=sys.stderr)
        return 1

    output_path = args.output or str(input_path.with_name(input_path.stem + '_styled.xlsx'))

    wb_src = load_workbook(str(input_path))
    wb_dst = Workbook()
    wb_dst.remove(wb_dst.active)

    # Copy all sheets with values and merged cells
    for sn in wb_src.sheetnames:
        ws_src = wb_src[sn]
        ws_dst = wb_dst.create_sheet(title=sn)
        mr = ws_src.max_row
        mc = ws_src.max_column
        for r in range(1, mr + 1):
            for c in range(1, mc + 1):
                ws_dst.cell(r, c).value = ws_src.cell(r, c).value
        for merged in ws_src.merged_cells.ranges:
            ws_dst.merge_cells(str(merged))
        for c in range(1, mc + 1):
            letter = get_column_letter(c)
            if ws_src.column_dimensions[letter].width:
                ws_dst.column_dimensions[letter].width = ws_src.column_dimensions[letter].width

    # Apply styling per sheet type
    date_grouped = {'21.29.31', '1.2.4', '9.10.12', '13.14.16', '17.18.19',
                    '22.23.25', '11.15.20', '26.3'}
    for sn in wb_dst.sheetnames:
        ws = wb_dst[sn]
        if sn == 'Summary':
            style_summary_sheet(ws)
        elif sn == 'Logs':
            style_logs_sheet(ws)
        elif sn in date_grouped or '.' in sn:
            style_date_grouped_sheet(ws)

    wb_dst.save(str(output_path))
    print(f'Styled workbook saved: {output_path}')
    print(f'Sheets: {wb_dst.sheetnames}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
