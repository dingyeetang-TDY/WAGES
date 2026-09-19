# -*- coding: utf-8 -*-
"""Tests for attendance calculator."""
import pytest
from attendance_calculator import parse_time, calculate_day_hours


def test_parse_time_valid():
    assert parse_time('08:50') == 530
    assert parse_time('18:46') == 1126
    assert parse_time('00:00') == 0
    assert parse_time('23:59') == 1439


def test_parse_time_invalid():
    assert parse_time('') is None
    assert parse_time('abc') is None
    assert parse_time('8:50') == 530  # single digit hour ok
    assert parse_time('25:00') is None
    assert parse_time('12:60') is None


def test_two_punches_all_work_no_break():
    result = calculate_day_hours(['08:50', '18:46'])
    assert result['count'] == 2
    assert not result['is_odd']
    assert result['work_minutes'] == 596
    assert result['break_minutes'] == 0
    assert len(result['work_segments']) == 1
    assert len(result['break_segments']) == 0
    assert result['work_segments'][0] == (530, 1126, 596)


def test_four_punches_with_one_break():
    # 08:30-12:00 work, 12:00-13:00 break, 13:00-17:30 work
    result = calculate_day_hours(['08:30', '12:00', '13:00', '17:30'])
    assert result['count'] == 4
    assert not result['is_odd']
    assert result['work_minutes'] == 480  # 3.5h + 4.5h
    assert result['break_minutes'] == 60
    assert len(result['work_segments']) == 2
    assert len(result['break_segments']) == 1
    assert result['work_segments'][0] == (510, 720, 210)
    assert result['work_segments'][1] == (780, 1050, 270)
    assert result['break_segments'][0] == (720, 780, 60)


def test_six_punches_with_two_breaks():
    result = calculate_day_hours(['08:00', '12:00', '13:00', '17:00', '18:00', '20:00'])
    assert result['count'] == 6
    assert not result['is_odd']
    # work: 08:00-12:00 (240) + 13:00-17:00 (240) + 18:00-20:00 (120) = 600
    assert result['work_minutes'] == 600
    # break: 12:00-13:00 (60) + 17:00-18:00 (60) = 120
    assert result['break_minutes'] == 120
    assert len(result['work_segments']) == 3
    assert len(result['break_segments']) == 2


def test_odd_punches_anomaly():
    result = calculate_day_hours(['18:48'])
    assert result['count'] == 1
    assert result['is_odd']
    assert result['work_minutes'] == 0
    assert result['break_minutes'] == 0
    assert result['work_segments'] == []
    assert result['break_segments'] == []


def test_unsorted_input_sorted_internally():
    result = calculate_day_hours(['18:46', '08:50'])
    assert result['work_segments'][0] == (530, 1126, 596)


def test_empty_punches():
    result = calculate_day_hours([])
    assert result['count'] == 0
    assert not result['is_odd']
    assert result['work_minutes'] == 0
    assert result['break_minutes'] == 0


def test_duplicate_and_invalid_entries():
    result = calculate_day_hours(['08:50', 'abc', '18:46', ''])
    assert result['count'] == 2
    assert result['work_minutes'] == 596


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
