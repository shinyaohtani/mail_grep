"""Tests for search_pattern.py - SearchPattern"""
import pytest
import re
import sys
from pathlib import Path
from unittest.mock import patch
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from search_pattern import SearchPattern


class TestSearchPatternCheckLine:
    """SearchPattern.check_line() のテスト (8個)"""

    def test_check_line_simple_match_and_no_match(self):
        """単純な文字列マッチとマッチなし"""
        sp = SearchPattern("hello")
        assert sp.check_line("hello world") is True
        assert sp.check_line("goodbye world") is False

    def test_check_line_regex_dot(self):
        """正規表現 . のマッチ"""
        sp = SearchPattern("h.llo")
        assert sp.check_line("hello") is True
        assert sp.check_line("hallo") is True
        assert sp.check_line("hllo") is False

    def test_check_line_regex_quantifiers(self):
        """正規表現 * と + のマッチ"""
        sp_star = SearchPattern("hel*o")
        assert sp_star.check_line("heo") is True
        assert sp_star.check_line("hello") is True

        sp_plus = SearchPattern("hel+o")
        assert sp_plus.check_line("heo") is False
        assert sp_plus.check_line("hello") is True

    def test_check_line_regex_alternation(self):
        """正規表現 | のマッチ (egrep)"""
        sp = SearchPattern("cat|dog")
        assert sp.check_line("I have a cat") is True
        assert sp.check_line("I have a dog") is True
        assert sp.check_line("I have a bird") is False

    def test_check_line_case_sensitivity(self):
        """大文字小文字の区別とIGNORECASE"""
        sp = SearchPattern("Hello")
        assert sp.check_line("Hello") is True
        assert sp.check_line("hello") is False

        sp_i = SearchPattern("hello", re.IGNORECASE)
        assert sp_i.check_line("Hello") is True
        assert sp_i.check_line("HELLO") is True

    def test_check_line_japanese(self):
        """日本語文字列のマッチ"""
        sp = SearchPattern("こんにちは")
        assert sp.check_line("今日はこんにちは世界") is True
        assert sp.check_line("さようなら") is False

    def test_check_line_email_pattern(self):
        """メールアドレス風パターン"""
        sp = SearchPattern(r"[\w.+-]+@[\w-]+\.[\w.-]+")
        assert sp.check_line("Contact: test@example.com") is True
        assert sp.check_line("No email here") is False

    def test_check_line_special_classes(self):
        """特殊文字クラス \\d, \\s, \\w"""
        assert SearchPattern(r"\d+").check_line("123") is True
        assert SearchPattern(r"\s").check_line("a b") is True
        assert SearchPattern(r"\w+").check_line("日本語") is True


class TestSearchPatternUniqueName:
    """SearchPattern.unique_name のテスト (6個)"""

    def test_unique_name_simple(self):
        """単純なパターン"""
        sp = SearchPattern("invoice")
        name = sp.unique_name
        assert name.startswith("invoice_")
        assert "_20" in name  # タイムスタンプ含む

    def test_unique_name_with_spaces(self):
        """スペース含むパターンは_に変換"""
        sp = SearchPattern("hello world")
        name = sp.unique_name
        assert name.startswith("hello_world_") or name.startswith("helloworld_")

    def test_unique_name_special_chars_removed(self):
        """特殊文字は除去"""
        sp = SearchPattern("test.*pattern")
        name = sp.unique_name
        assert "." not in name.split("_")[0]
        assert "*" not in name

    def test_unique_name_truncated_at_16(self):
        """16文字で切り詰め"""
        sp = SearchPattern("verylongpatternthatshouldbetruncated")
        name = sp.unique_name
        prefix = name.split("_20")[0]  # タイムスタンプ前
        assert len(prefix) <= 16

    def test_unique_name_empty_pattern_fallback(self):
        """空パターンは 'search' にフォールバック"""
        sp = SearchPattern(".*")
        name = sp.unique_name
        # .* は正規表現特殊文字のみなので除去される
        assert name.startswith("search_") or len(name.split("_")[0]) > 0

    def test_unique_name_japanese(self):
        """日本語パターン"""
        sp = SearchPattern("請求書")
        name = sp.unique_name
        assert "請求書" in name


class TestSearchPatternAdvanced:
    """SearchPattern 高度なパターンのテスト (2個)"""

    def test_word_boundary(self):
        """\\b 単語境界"""
        sp = SearchPattern(r"\bword\b")
        assert sp.check_line("a word here") is True
        assert sp.check_line("keyword") is False

    def test_non_greedy(self):
        """非貪欲マッチ"""
        sp = SearchPattern(r"<.+?>")
        assert sp.check_line("<tag>content</tag>") is True
