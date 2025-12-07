"""Tests for mail_message.py - MailMessage and internal classes"""
import pytest
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from mail_message import MailMessage, _MailBlob, _MailHeaders, _MailBody
from mail_profile import MailProfile
from search_pattern import SearchPattern

# テスト用emlxファイルのパス
CORPUS_DIR = Path(__file__).parent / "corpus"


class TestMailBlob:
    """_MailBlob のテスト (2個)"""

    def test_blob_bytes_and_message(self):
        """通常emlxのバイト読み込みとMessage取得"""
        blob = _MailBlob(CORPUS_DIR / "plain_utf8.emlx")
        data = blob.bytes()
        assert b"From:" in data
        assert b"Hello Bob" in data
        msg = blob.message()
        assert "alice@example.com" in str(msg["From"])

    def test_blob_prefixed_size(self):
        """サイズプレフィックス付きemlxの処理"""
        blob = _MailBlob(CORPUS_DIR / "prefixed_size.emlx")
        data = blob.bytes()
        assert not data.startswith(b"1234")
        assert b"From:" in data
        msg = blob.message()
        assert msg["Subject"] == "Sized message"


class TestMailHeaders:
    """_MailHeaders のテスト (6個)"""

    @pytest.fixture
    def plain_headers(self):
        blob = _MailBlob(CORPUS_DIR / "plain_utf8.emlx")
        return _MailHeaders(blob.message())

    @pytest.fixture
    def broken_headers(self):
        blob = _MailBlob(CORPUS_DIR / "broken_header.emlx")
        return _MailHeaders(blob.message())

    @pytest.fixture
    def no_date_headers(self):
        blob = _MailBlob(CORPUS_DIR / "no_date.emlx")
        return _MailHeaders(blob.message())

    def test_id_str_and_link(self, plain_headers):
        """Message-IDとリンクの取得"""
        assert "plain-utf8@example.com" in plain_headers.id_str()
        link = plain_headers.link()
        assert link.startswith("message:")

    def test_date_str_and_dt(self, plain_headers):
        """日付の取得"""
        assert plain_headers.date_str() == "2024-01-01 10:00:00"
        dt = plain_headers.date_dt()
        assert isinstance(dt, datetime)
        assert dt.year == 2024

    def test_date_broken(self, broken_headers):
        """不正な日付の処理"""
        assert isinstance(broken_headers.date_str(), str)
        assert broken_headers.date_dt() is None

    def test_date_missing(self, no_date_headers):
        """Dateヘッダがない場合"""
        assert no_date_headers.date_dt() is None
        assert no_date_headers.date_str() == ""

    def test_lines(self, plain_headers):
        """ヘッダ行リストの取得"""
        lines = plain_headers.lines()
        assert len(lines) == 4
        assert any("Subject:" in line for line in lines)

    def test_from_addr(self, plain_headers):
        """Fromアドレスの取得"""
        assert "alice@example.com" in plain_headers.from_addr()


class TestMailBody:
    """_MailBody のテスト (4個)"""

    def test_body_lines_plain(self):
        """プレーンテキストの本文行取得"""
        blob = _MailBlob(CORPUS_DIR / "plain_utf8.emlx")
        body = _MailBody(blob.message())
        lines = body.lines()
        assert len(lines) > 0
        assert any("Hello Bob" in line for line, _ in lines)

    def test_body_lines_html_only(self):
        """HTML onlyメールの本文行取得"""
        blob = _MailBlob(CORPUS_DIR / "html_only.emlx")
        body = _MailBody(blob.message())
        lines = body.lines()
        assert len(lines) > 0
        content_types = {ct for _, ct in lines}
        assert "text/html" in content_types or "text/html_textonly" in content_types

    def test_body_lines_multipart(self):
        """マルチパートメールの本文行取得"""
        blob = _MailBlob(CORPUS_DIR / "html_multipart.emlx")
        body = _MailBody(blob.message())
        lines = body.lines()
        content_types = {ct for _, ct in lines}
        assert len(content_types) > 1

    def test_body_lines_non_utf8(self):
        """Shift_JISエンコードの本文"""
        blob = _MailBlob(CORPUS_DIR / "non_utf8_body.emlx")
        body = _MailBody(blob.message())
        lines = body.lines()
        assert len(lines) > 0


class TestMailMessage:
    """MailMessage のテスト (6個)"""

    def test_key_profile_plain(self):
        """key_profile() の基本テスト"""
        msg = MailMessage(CORPUS_DIR / "plain_utf8.emlx")
        profile = msg.key_profile()
        assert isinstance(profile, MailProfile)
        assert profile.subj == "Plain UTF8 sample"
        assert "alice@example.com" in profile.from_addr

    def test_key_profile_no_date(self):
        """日付なしメールのプロファイル"""
        msg = MailMessage(CORPUS_DIR / "no_date.emlx")
        profile = msg.key_profile()
        assert profile.date_dt is None

    def test_extract_plain_match(self):
        """extract() - プレーン本文のマッチ"""
        msg = MailMessage(CORPUS_DIR / "plain_utf8.emlx")
        pattern = SearchPattern("Hello")
        matches = msg.extract(pattern)
        assert len(matches) > 0
        assert any("Hello" in line for _, line in matches)

    def test_extract_header_match(self):
        """extract() - ヘッダのマッチ"""
        msg = MailMessage(CORPUS_DIR / "plain_utf8.emlx")
        pattern = SearchPattern("alice@example.com")
        matches = msg.extract(pattern)
        assert any(parttype == "header" for parttype, _ in matches)

    def test_extract_no_match(self):
        """extract() - マッチなし"""
        msg = MailMessage(CORPUS_DIR / "plain_utf8.emlx")
        pattern = SearchPattern("NOTFOUNDXYZ12345")
        matches = msg.extract(pattern)
        assert len(matches) == 0

    def test_extract_multi_hit(self):
        """extract() - 複数行マッチ"""
        msg = MailMessage(CORPUS_DIR / "multi_hit.emlx")
        pattern = SearchPattern("[Ii]nvoice")
        matches = msg.extract(pattern)
        assert len(matches) >= 2
