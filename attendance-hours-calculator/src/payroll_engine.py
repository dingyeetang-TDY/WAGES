# -*- coding: utf-8 -*-
"""Payroll calculation engine for attendance data.

This module implements the company payroll rules:
  - full-time vs part-time
  - standard/basic hours
  - overtime (OT) and overall overtime (OOT)
  - public holiday handling
  - rest/break deductions
  - full-attendance bonus eligibility
  - part-time wage calculation
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import yaml

try:
    import holidays
except ImportError:  # pragma: no cover
    holidays = None


@dataclass
class DailyRecord:
    day: int
    start_minutes: Optional[float] = None  # minutes since midnight
    end_minutes: Optional[float] = None
    gross_minutes: float = 0.0    # in-company time
    break_minutes: float = 0.0    # actual break taken
    net_minutes: float = 0.0      # gross - deductible break
    is_public_holiday: bool = False
    late_minutes: float = 0.0
    early_minutes: float = 0.0
    is_odd_punch: bool = False


@dataclass
class EmployeePayroll:
    employee_name: str
    is_full_time: bool = True
    hourly_rate: Optional[float] = None
    standard_hours_per_day: float = 7.5
    standard_days_per_month: float = 26.0
    records: List[DailyRecord] = field(default_factory=list)

    # Totals
    total_gross_minutes: float = 0.0
    total_net_minutes: float = 0.0
    total_break_minutes: float = 0.0
    basic_minutes: float = 0.0
    ot_minutes: float = 0.0
    ph_work_minutes: float = 0.0
    ph_standard_minutes: float = 0.0
    ph_ot_minutes: float = 0.0
    oot_minutes: float = 0.0
    final_work_minutes: float = 0.0
    part_time_wage: float = 0.0

    # Counts
    work_days: int = 0
    late_days: int = 0
    early_leave_days: int = 0
    absent_days: int = 0
    odd_punch_days: int = 0
    full_attendance_eligible: bool = False


class PayrollEngine:
    def __init__(self, config_path: Optional[str] = None):
        if config_path is None:
            config_path = str(Path(__file__).with_name('payroll_config.yaml'))
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        self._holiday_cache: Dict[int, Any] = {}

    def _get_holidays(self, year: int):
        if year in self._holiday_cache:
            return self._holiday_cache[year]
        if holidays is None:
            raise RuntimeError('python-holidays is required. Run: pip install holidays')
        country = self.config['company']['country']
        state = self.config['company'].get('state') or None
        try:
            h = holidays.country_holidays(country, years=year, subdiv=state)
        except TypeError:
            h = holidays.country_holidays(country, years=year)
        self._holiday_cache[year] = h
        return h

    def is_public_holiday(self, date_obj: datetime.date) -> bool:
        h = self._get_holidays(date_obj.year)
        return date_obj in h

    def parse_time(self, value) -> Optional[float]:
        """Convert datetime.time or timedelta to minutes."""
        if value is None:
            return None
        if isinstance(value, datetime.time):
            return value.hour * 60 + value.minute + value.second / 60.0
        if isinstance(value, datetime.timedelta):
            return value.total_seconds() / 60.0
        return None

    def minutes_to_time_str(self, mins: Optional[float]) -> str:
        if mins is None:
            return ''
        sign = '-' if mins < 0 else ''
        mins = abs(mins)
        h = int(mins // 60)
        m = int(mins % 60)
        s = int(round((mins - int(mins)) * 60))
        return f'{sign}{h:02d}:{m:02d}:{s:02d}'

    def minutes_to_hours(self, mins: float) -> float:
        return mins / 60.0

    def compute_deductible_break(
        self,
        actual_break_minutes: float,
        standard_minutes: float,
        mode: str = 'exceed_only',
    ) -> float:
        """Return break minutes to deduct from gross work time."""
        if mode == 'all':
            return actual_break_minutes
        # exceed_only
        return max(0.0, actual_break_minutes - standard_minutes)

    def compute_daily_record(
        self,
        record: DailyRecord,
        standard_am_in: datetime.time,
        standard_pm_out: datetime.time,
    ) -> DailyRecord:
        """Enrich a DailyRecord with late/early/PH/net calculations."""
        am_in_mins = standard_am_in.hour * 60 + standard_am_in.minute
        pm_out_mins = standard_pm_out.hour * 60 + standard_pm_out.minute

        if record.start_minutes is not None and record.start_minutes > am_in_mins:
            record.late_minutes = record.start_minutes - am_in_mins
        if record.end_minutes is not None and record.end_minutes < pm_out_mins:
            record.early_minutes = pm_out_mins - record.end_minutes
        return record

    def compute_employee_payroll(
        self,
        employee_name: str,
        is_full_time: bool,
        records: List[DailyRecord],
        year: int,
        month: int,
        standard_hours_per_day: Optional[float] = None,
        standard_days_per_month: Optional[float] = None,
        hourly_rate: Optional[float] = None,
    ) -> EmployeePayroll:
        """Compute full payroll for one employee."""
        cfg_rest = self.config['rest']
        cfg_ph = self.config['public_holiday']
        cfg_part_time = self.config['part_time']
        cfg_full_time = self.config['full_time']
        cfg_schedule = self.config['schedule']

        std_hours_day = standard_hours_per_day or cfg_full_time['standard_hours_per_day']
        std_days_month = standard_days_per_month or cfg_full_time['standard_days_per_month'] or 26.0
        std_minutes_day = std_hours_day * 60
        basic_minutes = std_days_month * std_minutes_day

        # Schedule times
        am_in = datetime.datetime.strptime(cfg_schedule['am_in'], '%H:%M').time()
        pm_out = datetime.datetime.strptime(cfg_schedule['pm_out'], '%H:%M').time()

        payroll = EmployeePayroll(
            employee_name=employee_name,
            is_full_time=is_full_time,
            hourly_rate=hourly_rate or cfg_part_time['default_hourly_rate'],
            standard_hours_per_day=std_hours_day,
            standard_days_per_month=std_days_month,
            records=records,
            basic_minutes=basic_minutes,
        )

        for rec in records:
            rec = self.compute_daily_record(rec, am_in, pm_out)
            payroll.total_gross_minutes += rec.gross_minutes
            payroll.total_break_minutes += rec.break_minutes
            payroll.total_net_minutes += rec.net_minutes

            if rec.net_minutes > 0:
                payroll.work_days += 1

            if rec.late_minutes > 0:
                payroll.late_days += 1
            if rec.early_minutes > 0:
                payroll.early_leave_days += 1
            if rec.is_odd_punch:
                payroll.odd_punch_days += 1

            # Public holiday flag
            if rec.is_public_holiday and rec.net_minutes > 0:
                payroll.ph_work_minutes += rec.net_minutes
                payroll.ph_standard_minutes += std_minutes_day

        # Overtime for full-time
        if is_full_time:
            payroll.ot_minutes = payroll.total_net_minutes - payroll.basic_minutes
            if cfg_ph['full_time_ph_as_ot']:
                payroll.ph_ot_minutes = payroll.ph_work_minutes - payroll.ph_standard_minutes
                payroll.oot_minutes = payroll.ot_minutes + payroll.ph_ot_minutes
            else:
                payroll.oot_minutes = payroll.ot_minutes
            payroll.final_work_minutes = payroll.total_net_minutes
        else:
            # Part-time: final = total net + PH double
            payroll.final_work_minutes = payroll.total_net_minutes
            if cfg_part_time['public_holiday_double_pay']:
                payroll.final_work_minutes += payroll.ph_work_minutes
            payroll.part_time_wage = (payroll.final_work_minutes / 60.0) * payroll.hourly_rate

        # Full-attendance bonus eligibility
        cfg_bonus = self.config['full_attendance_bonus']
        if cfg_bonus['enabled']:
            payroll.full_attendance_eligible = True
            if cfg_bonus['no_late'] and payroll.late_days > 0:
                payroll.full_attendance_eligible = False
            if cfg_bonus['no_early_leave'] and payroll.early_leave_days > 0:
                payroll.full_attendance_eligible = False
            if cfg_bonus['no_absence'] and payroll.absent_days > 0:
                payroll.full_attendance_eligible = False
            if cfg_bonus['no_odd_punch'] and payroll.odd_punch_days > 0:
                payroll.full_attendance_eligible = False

        return payroll

    def format_payroll_report(self, payroll: EmployeePayroll) -> Dict[str, Any]:
        """Return a dictionary suitable for Excel/JSON export."""
        return {
            'employee_name': payroll.employee_name,
            'is_full_time': payroll.is_full_time,
            'work_days': payroll.work_days,
            'late_days': payroll.late_days,
            'early_leave_days': payroll.early_leave_days,
            'odd_punch_days': payroll.odd_punch_days,
            'full_attendance_eligible': payroll.full_attendance_eligible,
            'total_gross_hours': self.minutes_to_hours(payroll.total_gross_minutes),
            'total_net_hours': self.minutes_to_hours(payroll.total_net_minutes),
            'total_break_hours': self.minutes_to_hours(payroll.total_break_minutes),
            'basic_hours': self.minutes_to_hours(payroll.basic_minutes),
            'ot_hours': self.minutes_to_hours(payroll.ot_minutes),
            'ph_work_hours': self.minutes_to_hours(payroll.ph_work_minutes),
            'ph_standard_hours': self.minutes_to_hours(payroll.ph_standard_minutes),
            'ph_ot_hours': self.minutes_to_hours(payroll.ph_ot_minutes),
            'oot_hours': self.minutes_to_hours(payroll.oot_minutes),
            'final_work_hours': self.minutes_to_hours(payroll.final_work_minutes),
            'part_time_wage': round(payroll.part_time_wage, 2),
        }
