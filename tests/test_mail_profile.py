"""Tests for mail_profile.py - MailProfile"""
import pytest
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from mail_profile import MailProfile


class TestMailProfile:
    """MailProfile dataclass のテスト (8個)"""

    def test_create_mail_profile(self):
        """MailProfileの基本生成"""
        profile = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01 10:00:00",
            date_dt=datetime(2024, 1, 1, 10, 0, 0),
            link="message:%3Ctest@example.com%3E",
            subj="Test Subject",
            from_addr="sender@example.com",
            to_addr="receiver@example.com",
        )
        assert profile.message_id == "<test@example.com>"
        assert profile.date_str == "2024-01-01 10:00:00"
        assert profile.subj == "Test Subject"

    def test_mail_profile_immutable(self):
        """MailProfileは不変（frozen=True）"""
        profile = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01 10:00:00",
            date_dt=datetime(2024, 1, 1, 10, 0, 0),
            link="message:%3Ctest@example.com%3E",
            subj="Test Subject",
            from_addr="sender@example.com",
            to_addr="receiver@example.com",
        )
        with pytest.raises(Exception):  # FrozenInstanceError
            profile.subj = "New Subject"

    def test_excel_link_with_link(self):
        """excel_linkプロパティ - linkがある場合"""
        profile = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01 10:00:00",
            date_dt=datetime(2024, 1, 1, 10, 0, 0),
            link="message:%3Ctest@example.com%3E",
            subj="Test Subject",
            from_addr="sender@example.com",
            to_addr="receiver@example.com",
        )
        excel_link = profile.excel_link
        assert excel_link.startswith('=HYPERLINK("')
        assert "message:" in excel_link
        assert 'メール")' in excel_link

    def test_excel_link_without_link(self):
        """excel_linkプロパティ - linkが空の場合"""
        profile = MailProfile(
            message_id="",
            date_str="2024-01-01 10:00:00",
            date_dt=datetime(2024, 1, 1, 10, 0, 0),
            link="",
            subj="Test Subject",
            from_addr="sender@example.com",
            to_addr="receiver@example.com",
        )
        assert profile.excel_link == ""

    def test_mail_profile_date_dt_none(self):
        """date_dtがNoneの場合"""
        profile = MailProfile(
            message_id="<test@example.com>",
            date_str="",
            date_dt=None,
            link="message:%3Ctest@example.com%3E",
            subj="Test Subject",
            from_addr="sender@example.com",
            to_addr="receiver@example.com",
        )
        assert profile.date_dt is None
        assert profile.date_str == ""

    def test_mail_profile_japanese_subject(self):
        """日本語件名"""
        profile = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01 10:00:00",
            date_dt=datetime(2024, 1, 1, 10, 0, 0),
            link="message:%3Ctest@example.com%3E",
            subj="請求書のご確認",
            from_addr="sender@example.com",
            to_addr="receiver@example.com",
        )
        assert profile.subj == "請求書のご確認"

    def test_mail_profile_equality(self):
        """MailProfile同士の等価比較"""
        profile1 = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01",
            date_dt=None,
            link="",
            subj="Test",
            from_addr="a@a.com",
            to_addr="b@b.com",
        )
        profile2 = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01",
            date_dt=None,
            link="",
            subj="Test",
            from_addr="a@a.com",
            to_addr="b@b.com",
        )
        assert profile1 == profile2

    def test_mail_profile_hash(self):
        """MailProfileはハッシュ可能（frozen dataclass）"""
        profile = MailProfile(
            message_id="<test@example.com>",
            date_str="2024-01-01",
            date_dt=None,
            link="",
            subj="Test",
            from_addr="a@a.com",
            to_addr="b@b.com",
        )
        # setやdictのキーとして使える
        profile_set = {profile}
        assert profile in profile_set

    def test_mail_profile_empty_fields(self):
        """空フィールドを持つMailProfile"""
        profile = MailProfile(
            message_id="",
            date_str="",
            date_dt=None,
            link="",
            subj="",
            from_addr="",
            to_addr="",
        )
        assert profile.message_id == ""
        assert profile.excel_link == ""

    def test_mail_profile_special_characters(self):
        """特殊文字を含むMailProfile"""
        profile = MailProfile(
            message_id="<test+special@example.com>",
            date_str="2024-01-01",
            date_dt=None,
            link="message:%3Ctest%2Bspecial%40example.com%3E",
            subj="件名に「」や【】を含む",
            from_addr="user@例え.jp",
            to_addr="user@example.com",
        )
        assert "test+special" in profile.message_id
        assert "件名" in profile.subj
