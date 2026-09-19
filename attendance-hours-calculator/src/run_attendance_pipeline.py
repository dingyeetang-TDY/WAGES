# -*- coding: utf-8 -*-
"""One-stop pipeline for attendance processing.

Workflow:
    原始 .xls → 转换矩阵 → 标注奇数 → 手动修正 → 刷新标注 → 计算工时

Usage:
    # Step 1~2: convert & highlight odd punches (produces matrix for manual fix)
    python run_attendance_pipeline.py input.xls

    # After you manually correct the highlighted cells, run step 4~5:
    # (automatically refreshes red highlighting based on current values, then calculates)
    python run_attendance_pipeline.py --calculate corrected_matrix.xlsx --header-row 2
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
from typing import Optional, Sequence

from convert_xls import convert as convert_xls
from highlight_odd_punches import highlight_odd_punches


SCRIPT_DIR = Path(__file__).resolve().parent
CALCULATOR = SCRIPT_DIR / 'attendance_calculator.py'


def step_convert_highlight(input_path: str, output_path: Optional[str] = None) -> tuple:
    """Convert original .xls to matrix and highlight odd punches."""
    print('=' * 60)
    print('Step 1/5: Converting original .xls to matrix format...')
    print('=' * 60)
    matrix_path = convert_xls(input_path, output_path)
    print()

    print('=' * 60)
    print('Step 2/5: Highlighting odd-punch cells...')
    print('=' * 60)
    highlighted_path, highlighted, date_count = highlight_odd_punches(matrix_path)
    print()

    return highlighted_path, highlighted, date_count


def step_calculate(corrected_path: str, header_row: int = 1, output_path: Optional[str] = None) -> int:
    """Refresh highlighting on corrected matrix, then run attendance_calculator.py."""
    print('=' * 60)
    print('Step 4/5: Refreshing odd-punch highlighting based on current values...')
    print('=' * 60)
    refreshed_path, highlighted, date_count = highlight_odd_punches(corrected_path)
    print(f'Refreshed file: {refreshed_path}')
    print(f'Date columns: {date_count}')
    print(f'Remaining odd-punch cells: {highlighted}')
    print()

    print('=' * 60)
    print('Step 5/5: Calculating work hours from corrected matrix...')
    print('=' * 60)

    cmd = [
        sys.executable,
        str(CALCULATOR),
        refreshed_path,
        '--header-row', str(header_row),
    ]
    if output_path:
        cmd.extend(['-o', output_path])

    result = subprocess.run(cmd, check=False)
    return result.returncode


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description='One-stop attendance pipeline: convert → highlight → (manual fix) → calculate'
    )
    parser.add_argument('input', help='Input file path')
    parser.add_argument('--calculate', action='store_true',
                        help='Run in calculation mode on an already-corrected matrix xlsx')
    parser.add_argument('--header-row', type=int, default=2,
                        help='Header row in corrected matrix (1-based; default 2 because of the note row)')
    parser.add_argument('-o', '--output', default=None, help='Output file path')
    args = parser.parse_args(argv)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'Error: file not found: {input_path}', file=sys.stderr)
        return 1

    if args.calculate:
        # Mode: calculate from corrected matrix
        return step_calculate(str(input_path), header_row=args.header_row, output_path=args.output)

    # Mode: convert + highlight
    highlighted_path, highlighted, date_count = step_convert_highlight(str(input_path), args.output)

    print('=' * 60)
    print('Pipeline complete. Next step: manual correction (Step 3/5)')
    print('=' * 60)
    print(f'Output file: {highlighted_path}')
    print(f'Date columns: {date_count}')
    print(f'Odd-punch cells highlighted: {highlighted}')
    print()
    print('Please:')
    print('  1. Open the output file in Excel.')
    print('  2. Fix every RED cell so it contains an EVEN number of HH:MM times.')
    print('  3. Save the corrected file.')
    print()
    print('Then run (will auto-refresh highlighting before calculating):')
    print(f'  python {Path(__file__).name} "{highlighted_path}" --calculate --header-row 2')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
