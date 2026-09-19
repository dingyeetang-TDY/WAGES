#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate an HTML attendance report from a pre-parsed JSON file.

This is the second half of a two-step pipeline:
    Step 1: parse_logs.py  (xlsx/xls → JSON)
    Step 2: generate_report_from_json.py  (JSON → HTML)

For a one-step alternative, use generate_report.py directly.

Usage:
    python generate_report_from_json.py parsed.json
    python generate_report_from_json.py parsed.json -o report.html
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def fmt_hours(mins: float) -> str:
    if mins <= 0:
        return '0.00'
    return f'{mins / 60:.2f}'


def fmt_hm(mins: float) -> str:
    if mins <= 0:
        return '0h00m'
    h = int(mins) // 60
    m = int(mins) % 60
    return f'{h}h{m:02d}m'


def generate_html(employees: list, dates: list, input_name: str) -> str:
    total_employees = len(employees)
    total_present_days = sum(
        1 for e in employees for r in e['records'].values() if len(r['times']) > 0
    )
    total_odd_days = sum(len(e['odd_annotations']) for e in employees)
    total_work_mins = sum(sum(e['daily_work_mins'].values()) for e in employees)

    annotated_emps = [
        (e['name'], e['no'], e['dept'], e['odd_annotations'])
        for e in employees if e['odd_annotations']
    ]
    annotated_emps.sort(key=lambda x: int(x[1]) if x[1].isdigit() else 999)
    employees.sort(key=lambda e: int(e['no']) if e['no'].isdigit() else 999)

    period = f'{dates[0]} ~ {dates[-1]}' if dates else ''

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>员工打卡工时统计报表 ({period})</title>
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family:'Microsoft YaHei','Noto Sans CJK SC',sans-serif; background:#f0f2f5; color:#1a1a2e; padding:20px; }}
  .container {{ max-width:1600px; margin:0 auto; }}
  h1 {{ text-align:center; font-size:24px; margin-bottom:5px; color:#16213e; }}
  .subtitle {{ text-align:center; color:#666; margin-bottom:20px; font-size:14px; }}
  .section {{ background:#fff; border-radius:10px; padding:20px; margin-bottom:25px; box-shadow:0 2px 12px rgba(0,0,0,0.08); }}
  .section h2 {{ font-size:18px; color:#16213e; margin-bottom:15px; padding-bottom:10px; border-bottom:2px solid #e8ecf1; }}
  .section h3 {{ font-size:15px; color:#2c3e50; margin:15px 0 10px 0; }}
  .stats-row {{ display:flex; gap:15px; flex-wrap:wrap; margin-bottom:20px; }}
  .stat-card {{ flex:1; min-width:150px; background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); color:#fff; border-radius:10px; padding:15px; text-align:center; }}
  .stat-card.alert {{ background:linear-gradient(135deg,#f093fb 0%,#f5576c 100%); }}
  .stat-card.success {{ background:linear-gradient(135deg,#4facfe 0%,#00f2fe 100%); }}
  .stat-card.warn {{ background:linear-gradient(135deg,#fa709a 0%,#fee140 100%); }}
  .stat-card .num {{ font-size:28px; font-weight:bold; }}
  .stat-card .label {{ font-size:12px; opacity:0.9; margin-top:4px; }}
  .legend {{ display:flex; gap:20px; flex-wrap:wrap; margin:15px 0; padding:10px; background:#f8f9fa; border-radius:8px; }}
  .legend-item {{ display:flex; align-items:center; gap:6px; font-size:13px; }}
  .legend-box {{ width:20px; height:16px; border-radius:3px; border:1px solid #ddd; }}
  .annotation-list {{ margin-top:10px; }}
  .annotation-item {{ padding:10px 15px; margin:5px 0; border-radius:8px; border-left:4px solid; }}
  .annotation-item.odd-item {{ border-left-color:#8B5CF6; background:#F3E8FF; }}
  .annotation-item .emp-name {{ font-weight:bold; font-size:14px; }}
  .annotation-item .emp-info {{ font-size:12px; color:#666; }}
  .annotation-item .detail {{ font-size:13px; margin-top:5px; }}
  .annotation-item .time-tag {{ font-weight:bold; font-family:monospace; }}
  .emp-block {{ margin-bottom:30px; padding-bottom:20px; border-bottom:1px dashed #ddd; }}
  .emp-block:last-child {{ border-bottom:none; }}
  .emp-header {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap; gap:10px; }}
  .emp-title {{ font-size:16px; font-weight:bold; color:#1e3a5f; }}
  .emp-meta {{ font-size:12px; color:#888; }}
  .day-grid {{ display:grid; grid-template-columns:repeat(auto-fill, minmax(280px, 1fr)); gap:12px; }}
  .day-card {{ border-radius:8px; border:1px solid #e5e7eb; overflow:hidden; }}
  .day-card-header {{ padding:8px 12px; font-size:13px; font-weight:bold; display:flex; justify-content:space-between; align-items:center; }}
  .day-card-body {{ padding:10px 12px; font-size:13px; }}
  .day-card.normal .day-card-header {{ background:#EFF6FF; color:#1E40AF; }}
  .day-card.normal {{ background:#fff; }}
  .day-card.odd .day-card-header {{ background:#F3E8FF; color:#6B21A8; }}
  .day-card.odd {{ background:#FAF5FF; border-color:#C4B5FD; }}
  .day-card.empty .day-card-header {{ background:#F3F4F6; color:#6B7280; }}
  .day-card.empty {{ background:#F9FAFB; }}
  .times-list {{ display:flex; flex-wrap:wrap; gap:4px; margin-bottom:8px; }}
  .time-pill {{ padding:2px 8px; border-radius:12px; font-size:12px; font-family:monospace; background:#E5E7EB; color:#374151; }}
  .time-pill.odd-mark {{ background:#F3E8FF; color:#6B21A8; font-weight:bold; border:1px solid #C4B5FD; }}
  .segments {{ margin-top:6px; }}
  .seg-row {{ display:flex; align-items:center; gap:6px; margin:3px 0; font-size:12px; }}
  .seg-badge {{ padding:1px 6px; border-radius:4px; font-size:11px; font-weight:bold; white-space:nowrap; }}
  .seg-work {{ background:#D1FAE5; color:#065F46; }}
  .seg-rest {{ background:#FEF3C7; color:#92400E; }}
  .seg-rest2 {{ background:#FEE2E2; color:#991B1B; }}
  .seg-time {{ font-family:monospace; color:#4B5563; }}
  .seg-dur {{ color:#6B7280; font-size:11px; }}
  .day-summary {{ margin-top:8px; padding-top:6px; border-top:1px dashed #E5E7EB; font-size:13px; font-weight:bold; color:#1E40AF; }}
  .day-summary .work-only {{ color:#059669; }}
  .day-summary .total-only {{ color:#6B7280; font-weight:normal; font-size:11px; }}
  .odd-banner {{ background:#8B5CF6; color:#fff; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:bold; }}
  .summary-table {{ width:100%; border-collapse:collapse; font-size:13px; margin-top:15px; }}
  .summary-table th {{ background:#2c3e50; color:#fff; padding:8px 6px; text-align:center; }}
  .summary-table td {{ padding:6px 4px; text-align:center; border-bottom:1px solid #eee; }}
  .summary-table tr:nth-child(even) {{ background:#f8f9fa; }}
  .summary-table td.left {{ text-align:left; padding-left:10px; }}
  .summary-table .total-row {{ font-weight:bold; background:#e3f2fd !important; }}
  .summary-table .odd-cell {{ background:#F3E8FF; color:#6B21A8; font-weight:bold; }}
  .rule-box {{ background:#F0F9FF; border-left:4px solid #0EA5E9; padding:12px 16px; border-radius:0 8px 8px 0; margin:10px 0; font-size:13px; color:#0C4A6E; }}
  .rule-box strong {{ color:#0369A1; }}
</style>
</head>
<body>
<div class="container">
<h1>员工打卡工时统计报表</h1>
<p class="subtitle">考勤周期: {period} | 打卡规则 V2 | 数据源: {input_name}</p>

<div class="stats-row">
  <div class="stat-card"><div class="num">{total_employees}</div><div class="label">员工总数</div></div>
  <div class="stat-card success"><div class="num">{len(dates)}</div><div class="label">考勤天数</div></div>
  <div class="stat-card"><div class="num">{total_present_days}</div><div class="label">总出勤人次</div></div>
  <div class="stat-card alert"><div class="num">{total_odd_days}</div><div class="label">奇数打卡异常</div></div>
  <div class="stat-card warn"><div class="num">{fmt_hours(total_work_mins)}</div><div class="label">总工时(仅工作时段)</div></div>
</div>

<div class="section">
<h2>工时计算规则</h2>
<div class="rule-box">
  <strong>规则一：</strong>如果某日的打卡时间数量<strong>不是偶数</strong>（即为奇数），则标注为异常打卡。<br>
  <strong>规则二：</strong>工时仅计算工作时段，中场休息时间不计入工时：<br>
  &nbsp;&nbsp;&bull; 2个打卡时间 &rarr; 全部算工时<br>
  &nbsp;&nbsp;&bull; 4个打卡时间 &rarr; 第1-2段=上午班(算)，第2-3段=中场休息(不算)，第3-4段=下午班(算)<br>
  &nbsp;&nbsp;&bull; 6个打卡时间 &rarr; 上午班(算) + 中场休息(不算) + 下午班(算) + 中场休息2(不算) + 下午班2(算)
</div>
<div class="legend">
  <div class="legend-item"><div class="legend-box" style="background:#D1FAE5;"></div> 工作时段（计入工时）</div>
  <div class="legend-item"><div class="legend-box" style="background:#FEF3C7;"></div> 中场休息（不计工时）</div>
  <div class="legend-item"><div class="legend-box" style="background:#F3E8FF;"></div> 奇数打卡异常</div>
  <div class="legend-item"><div class="legend-box" style="background:#F3F4F6;"></div> 无打卡记录</div>
</div>
'''

    if annotated_emps:
        html += '<h3>奇数打卡异常汇总</h3><div class="annotation-list">'
        for name, no, dept, odd_dict in annotated_emps:
            details = ', '.join(
                f'{d} <span class="time-tag">{cnt}次打卡</span>' for d, cnt in odd_dict.items()
            )
            html += f'''<div class="annotation-item odd-item">
<span class="emp-name">{name}</span> <span class="emp-info">(No.{no} | {dept})</span>
<div class="detail">{details}</div></div>'''
        html += '</div>'
    else:
        html += '<p style="color:#28a745;text-align:center;padding:20px;">本期无奇数打卡异常</p>'
    html += '</div>'

    html += '<div class="section"><h2>各员工每日打卡明细与工时</h2>'
    for emp in employees:
        odd_dict = emp['odd_annotations']
        emp_total_work = sum(emp['daily_work_mins'].values())
        html += f'''<div class="emp-block">
<div class="emp-header">
<div class="emp-title">{emp['name']} <span style="font-size:13px;color:#666;">(No.{emp['no']} | {emp['dept']})</span></div>
<div class="emp-meta">总工时(工作时段): <strong>{fmt_hours(emp_total_work)}h</strong> | 异常天数: <strong style="color:#8B5CF6;">{len(odd_dict)}天</strong></div>
</div>
<div class="day-grid">'''
        for date_str in dates:
            rec = emp['records'].get(date_str, {'times': [], 'status': 'No Data'})
            times = rec['times']
            n = len(times)
            segments = emp['segments'].get(date_str, [])
            work_mins = emp['daily_work_mins'].get(date_str, 0)

            if n == 0:
                card_class = 'empty'
                header_extra = ''
                body_content = '<div style="color:#9CA3AF;font-size:12px;text-align:center;padding:10px;">无打卡记录</div>'
            else:
                if n % 2 != 0:
                    card_class = 'odd'
                    header_extra = f'<span class="odd-banner">{n}次打卡(奇数)</span>'
                else:
                    card_class = 'normal'
                    header_extra = f'<span style="font-size:11px;color:#6B7280;">{n}次打卡</span>'

                raw_mins = sum(seg[3] for seg in segments)
                times_html = '<div class="times-list">'
                for t in times:
                    cls = 'time-pill odd-mark' if n % 2 != 0 else 'time-pill'
                    times_html += f'<span class="{cls}">{t}</span>'
                times_html += '</div>'

                seg_html = '<div class="segments">'
                for seg in segments:
                    start, end, seg_type, seg_mins = seg
                    if '休息' in seg_type:
                        badge_cls = 'seg-rest2' if '2' in seg_type else 'seg-rest'
                    else:
                        badge_cls = 'seg-work'
                    seg_html += f'''<div class="seg-row">
<span class="seg-badge {badge_cls}">{seg_type}</span>
<span class="seg-time">{start} &rarr; {end}</span>
<span class="seg-dur">{fmt_hm(seg_mins)}</span>
</div>'''
                seg_html += '</div>'

                body_content = times_html + seg_html
                body_content += f'''<div class="day-summary">
工作时段工时: <span class="work-only">{fmt_hours(work_mins)}h</span>
<span class="total-only">(全时段总时长: {fmt_hours(raw_mins)}h)</span>
</div>'''

            html += f'''<div class="day-card {card_class}">
<div class="day-card-header"><span>{date_str}</span>{header_extra}</div>
<div class="day-card-body">{body_content}</div>
</div>'''
        html += '</div></div>'
    html += '</div>'

    html += '''<div class="section">
<h2>工时汇总表</h2>
<table class="summary-table">
<thead><tr>
<th>No.</th><th>姓名</th><th>部门</th><th>有打卡天数</th><th>总工时(工作时段)</th><th>总时长(含休息)</th><th>有效占比</th><th>奇数异常天数</th>
</tr></thead><tbody>'''
    grand_work = 0
    grand_raw = 0
    for emp in employees:
        work = sum(emp['daily_work_mins'].values())
        raw = 0
        present = 0
        for d, rec in emp['records'].items():
            if len(rec['times']) > 0:
                present += 1
                for seg in emp['segments'].get(d, []):
                    raw += seg[3]
        grand_work += work
        grand_raw += raw
        ratio = f'{work / raw * 100:.1f}%' if raw > 0 else '--'
        odd_cnt = len(emp['odd_annotations'])
        odd_cell = f'<span class="odd-cell">{odd_cnt}</span>' if odd_cnt > 0 else '0'
        html += f'''<tr>
<td>{emp['no']}</td><td class="left">{emp['name']}</td><td class="left">{emp['dept']}</td>
<td>{present}</td><td><strong>{fmt_hours(work)}h</strong></td><td>{fmt_hours(raw)}h</td>
<td>{ratio}</td><td>{odd_cell}</td></tr>'''

    grand_ratio = f'{grand_work / grand_raw * 100:.1f}%' if grand_raw > 0 else '--'
    html += f'''<tr class="total-row">
<td colspan="3" class="left">合计</td><td>--</td>
<td><strong>{fmt_hours(grand_work)}h</strong></td><td>{fmt_hours(grand_raw)}h</td>
<td>{grand_ratio}</td><td>{total_odd_days}</td>
</tr></tbody></table></div></div></body></html>'''
    return html


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description='Generate an HTML attendance report from a pre-parsed JSON file.'
    )
    parser.add_argument('input', help='Input .json file path (from parse_logs.py)')
    parser.add_argument('-o', '--output', default=None, help='Output .html file path')
    args = parser.parse_args(argv)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f'Error: file not found: {input_path}', file=sys.stderr)
        return 1

    output_path = args.output or str(input_path.with_suffix('.html'))

    with open(args.input, 'r', encoding='utf-8') as f:
        employees = json.load(f)

    # Derive dates from the first employee's records
    dates = list(employees[0]['records'].keys()) if employees else []

    html = generate_html(employees, dates, input_path.name)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    total_odd = sum(len(e['odd_annotations']) for e in employees)
    total_work = sum(sum(e['daily_work_mins'].values()) for e in employees)
    print(f'HTML report saved: {output_path}')
    print(f'Employees: {len(employees)}')
    print(f'Odd-punch anomaly days: {total_odd}')
    print(f'Total work hours: {fmt_hours(total_work)}h')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
