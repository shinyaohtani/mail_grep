"""Tests for hit_line.py - HitLine"""
import pytest
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from hit_line import HitLine
from mail_profile import MailProfile


@pytest.fixture
def sample_profile():
    """テスト用MailProfile"""
    return MailProfile(
        message_id="<test@example.com>",
        date_str="2024-01-01 10:00:00",
        date_dt=datetime(2024, 1, 1, 10, 0, 0),
        link="message:%3Ctest@example.com%3E",
        subj="Test Subject",
        from_addr="sender@example.com",
        to_addr="receiver@example.com",
    )


class TestHitLine:
    """HitLine のテスト (8個)"""

    def test_create_hit_line(self, sample_profile):
        """HitLineの基本生成"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="text/plain",
            line="This is a matched line.",
        )
        assert hit.mail_id == 1
        assert hit.parttype == "text/plain"
        assert hit.line == "This is a matched line."

    def test_hit_line_strips_whitespace(self, sample_profile):
        """行の前後空白は除去"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="text/plain",
            line="  padded line  ",
        )
        assert hit.line == "padded line"

    def test_hit_line_values_count_and_order(self, sample_profile):
        """values()の要素数と順序"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=5,
            hit_count=3,
            parttype="text/plain",
            line="Content",
        )
        values = hit.values()
        assert len(values) == 10
        assert values[0] == "5"  # mail_id
        assert values[1] == "3"  # hit_count

    def test_hit_line_values_sanitized(self, sample_profile):
        """values()は改行を⏎に変換"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="text/plain",
            line="Line with\nnewline",
        )
        values = hit.values()
        assert "\n" not in values[-1]
        assert "⏎" in values[-1]

    def test_hit_line_values_control_chars_removed(self, sample_profile):
        """values()で制御文字が除去される"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="text/plain",
            line="Line with\x00null",
        )
        values = hit.values()
        assert "\x00" not in values[-1]

    def test_hit_line_japanese_content(self, sample_profile):
        """日本語コンテンツの処理"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="text/plain",
            line="請求書の確認をお願いします",
        )
        values = hit.values()
        assert "請求書" in values[-1]

    def test_hit_line_header_parttype(self, sample_profile):
        """headerパートタイプ"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="header",
            line="From: sender@example.com",
        )
        assert hit.parttype == "header"
        values = hit.values()
        assert values[8] == "header"

    def test_hit_line_empty_line(self, sample_profile):
        """空行の処理"""
        hit = HitLine(
            mail_keys=sample_profile,
            mail_id=1,
            hit_count=1,
            parttype="text/plain",
            line="   ",
        )
        assert hit.line == ""
