"""Tests for smart_logging.py - SmartLogging"""
import pytest
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from smart_logging import SmartLogging, trancate, PROJ_ABSPATH


class TestTruncate:
    """trancate() 関数のテスト (2個)"""

    def test_trancate_short_and_exact(self):
        """短いテキストとちょうど最大長"""
        assert trancate("hello", 10) == "hello"
        assert trancate("hello", 5) == "hello"

    def test_trancate_long(self):
        """長いテキストは切り詰め"""
        result = trancate("hello world", 8)
        assert result == "hello..."
        assert len(result) == 8


class TestSmartLoggingBasic:
    """SmartLogging 基本機能のテスト (3個)"""

    def test_create_smart_logging(self):
        """SmartLoggingの生成"""
        sl = SmartLogging()
        assert sl._level == logging.INFO
        sl_debug = SmartLogging(logging.DEBUG)
        assert sl_debug._level == logging.DEBUG

    def test_context_manager(self):
        """with文での使用"""
        logger = logging.getLogger()
        original_handlers = logger.handlers.copy()
        logger.handlers = []

        try:
            with SmartLogging() as sl:
                assert sl._stream_handler is not None
        finally:
            logger.handlers = original_handlers

    def test_proj_abspath_exists(self):
        """PROJ_ABSPATHが正しく設定されている"""
        assert PROJ_ABSPATH.exists()
        assert PROJ_ABSPATH.is_dir()


class TestSmartLoggingStreamControl:
    """SmartLogging ストリーム制御のテスト (2個)"""

    def test_set_stream_level(self):
        """set_stream_level()"""
        logger = logging.getLogger()
        original_handlers = logger.handlers.copy()
        logger.handlers = []

        try:
            with SmartLogging() as sl:
                sl.set_stream_level(logging.DEBUG)
                assert sl._stream_handler.level == logging.DEBUG
                sl.set_stream_level(logging.WARNING)
                assert sl._stream_handler.level == logging.WARNING
        finally:
            logger.handlers = original_handlers

    def test_set_stream_filter(self):
        """set_stream_filter() on/off"""
        logger = logging.getLogger()
        original_handlers = logger.handlers.copy()
        logger.handlers = []

        try:
            with SmartLogging() as sl:
                sl.set_stream_filter(True)
                assert any(
                    isinstance(f, SmartLogging._OnlyMyLogsFilter)
                    for f in sl._stream_handler.filters
                )
                sl.set_stream_filter(False)
                assert not any(
                    isinstance(f, SmartLogging._OnlyMyLogsFilter)
                    for f in sl._stream_handler.filters
                )
        finally:
            logger.handlers = original_handlers


class TestOnlyMyLogsFilter:
    """SmartLogging._OnlyMyLogsFilter のテスト (2個)"""

    def test_filter_project_file(self):
        """プロジェクト内ファイルは通過"""
        filter_obj = SmartLogging._OnlyMyLogsFilter()
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname=str(PROJ_ABSPATH / "mail_grep.py"),
            lineno=1,
            msg="test",
            args=(),
            exc_info=None,
        )
        assert filter_obj.filter(record) is True

    def test_filter_external_file(self):
        """プロジェクト外ファイルはブロック"""
        filter_obj = SmartLogging._OnlyMyLogsFilter()
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="/usr/lib/python/logging.py",
            lineno=1,
            msg="test",
            args=(),
            exc_info=None,
        )
        assert filter_obj.filter(record) is False
