"""Tests for mail_string_utils.py - CsvFieldText, AnyText, EncodedHeader"""
import pytest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from mail_string_utils import CsvFieldText, AnyText, EncodedHeader


class TestCsvFieldText:
    """CsvFieldText.sanitize() のテスト (8個)"""

    def test_sanitize_plain_string(self):
        """通常文字列はそのまま返る"""
        assert CsvFieldText.sanitize("hello world") == "hello world"

    def test_sanitize_none_returns_empty(self):
        """Noneは空文字列"""
        assert CsvFieldText.sanitize(None) == ""

    def test_sanitize_number_types(self):
        """数値型は文字列化される"""
        assert CsvFieldText.sanitize(123) == "123"
        assert CsvFieldText.sanitize(3.14) == "3.14"

    def test_sanitize_newlines(self):
        """改行処理: LFは⏎に、CRは除去"""
        assert CsvFieldText.sanitize("line1\nline2") == "line1⏎line2"
        assert CsvFieldText.sanitize("line1\rline2") == "line1line2"
        assert CsvFieldText.sanitize("line1\r\nline2") == "line1⏎line2"

    def test_sanitize_control_chars_removed(self):
        """制御文字は除去"""
        assert CsvFieldText.sanitize("a\x00b") == "ab"
        assert CsvFieldText.sanitize("a\tb") == "ab"
        assert CsvFieldText.sanitize("a\x07b") == "ab"
        assert CsvFieldText.sanitize("a\x7fb") == "ab"

    def test_sanitize_multiple_control_chars(self):
        """複数の制御文字を一度に除去"""
        assert CsvFieldText.sanitize("\x00\x01hello\x1f\x7f") == "hello"

    def test_sanitize_japanese_preserved(self):
        """日本語は保持される"""
        assert CsvFieldText.sanitize("こんにちは") == "こんにちは"

    def test_sanitize_mixed_content(self):
        """混合コンテンツのサニタイズ"""
        assert CsvFieldText.sanitize("Hello\n世界\x00!") == "Hello⏎世界!"


class TestAnyText:
    """AnyText.to_str() のテスト (5個)"""

    def test_to_str_none_and_empty(self):
        """Noneと空は空文字列"""
        assert AnyText.to_str(None) == ""
        assert AnyText.to_str("") == ""
        assert AnyText.to_str(b"") == ""

    def test_to_str_string_passthrough(self):
        """文字列はそのまま"""
        assert AnyText.to_str("hello") == "hello"

    def test_to_str_bytes_utf8(self):
        """UTF-8バイトはデコードされる"""
        assert AnyText.to_str(b"hello") == "hello"
        assert AnyText.to_str("こんにちは".encode("utf-8")) == "こんにちは"

    def test_to_str_bytes_invalid_utf8_replaced(self):
        """不正UTF-8はreplaceで処理"""
        result = AnyText.to_str(b"\xff\xfe")
        assert "�" in result or result != ""

    def test_to_str_non_string_type_converted(self):
        """その他の型はstr()で変換"""
        assert AnyText.to_str(123) == "123"


class TestEncodedHeader:
    """EncodedHeader.decode() のテスト (7個)"""

    def test_decode_none_and_empty(self):
        """Noneと空文字列"""
        assert EncodedHeader.decode(None) == ""
        assert EncodedHeader.decode("") == ""

    def test_decode_plain_text(self):
        """プレーンテキストはそのまま"""
        assert EncodedHeader.decode("Hello World") == "Hello World"
        assert EncodedHeader.decode("こんにちは") == "こんにちは"

    def test_decode_utf8_base64(self):
        """UTF-8 Base64エンコード"""
        result = EncodedHeader.decode("=?UTF-8?B?44GT44KT44Gr44Gh44Gv?=")
        assert result == "こんにちは"

    def test_decode_utf8_quoted_printable(self):
        """UTF-8 Quoted-Printableエンコード"""
        result = EncodedHeader.decode("=?UTF-8?Q?=E3=83=86=E3=82=B9=E3=83=88?=")
        assert result == "テスト"

    def test_decode_iso2022jp_base64(self):
        """ISO-2022-JP Base64エンコード"""
        result = EncodedHeader.decode("=?ISO-2022-JP?B?GyRCJDMkcyRLJEEkTxsoQg==?=")
        assert result == "こんにちは"

    def test_decode_newlines(self):
        """改行処理: CRは除去、LFはスペースに"""
        assert "\r" not in EncodedHeader.decode("hello\rworld")
        assert EncodedHeader.decode("hello\nworld") == "hello world"
        assert EncodedHeader.decode("hello\r\nworld") == "hello world"

    def test_decode_invalid_encoding_fallback(self):
        """不正エンコーディングは例外を出さない"""
        result = EncodedHeader.decode("=?INVALID?B?broken?=")
        assert isinstance(result, str)
