"""Tests for mail_grep.py - MailGrepApp, AppArguments, and integration"""
import pytest
import sys
from pathlib import Path
import tempfile
import csv

sys.path.insert(0, str(Path(__file__).parent.parent))

from mail_grep import MailGrepApp, AppArguments
from mail_folder import MailFolder
from search_pattern import SearchPattern

CORPUS_DIR = Path(__file__).parent / "corpus"


class TestAppArguments:
    """AppArguments のテスト (4個)"""

    def test_parse_pattern_only(self):
        """パターンのみの引数解析"""
        sys.argv = ["mail_grep.py", "test pattern"]
        app_args = AppArguments()
        args = app_args.parse()
        assert args.pattern == "test pattern"

    def test_parse_options(self):
        """各種オプション"""
        sys.argv = ["mail_grep.py", "-i", "--only-sent", "test"]
        app_args = AppArguments()
        args = app_args.parse()
        assert args.ignore_case is True
        assert args.only_sent is True

    def test_parse_output_and_source(self):
        """-o, -s オプション"""
        sys.argv = ["mail_grep.py", "-o", "/tmp/out.csv", "-s", "/custom/path", "test"]
        app_args = AppArguments()
        args = app_args.parse()
        assert args.output == Path("/tmp/out.csv")
        assert args.source == Path("/custom/path")

    def test_parse_default_source(self):
        """デフォルトのsourceパス"""
        sys.argv = ["mail_grep.py", "test"]
        app_args = AppArguments()
        args = app_args.parse()
        expected = Path.home() / "Library" / "Mail" / "V10"
        assert args.source == expected


class TestMailGrepAppLinePreview:
    """MailGrepApp.line_preview() のテスト (2個)"""

    def test_line_preview_short(self):
        """短い行はそのまま"""
        preview = MailGrepApp.line_preview("short text")
        assert preview == "short text"

    def test_line_preview_long_truncated(self):
        """長い行は切り詰め"""
        long_text = "x" * 100
        preview = MailGrepApp.line_preview(long_text)
        assert len(preview) <= 53
        assert preview.endswith("...")


class TestMailGrepAppIntegration:
    """MailGrepApp 統合テスト (4個)"""

    def test_run_with_corpus(self):
        """corpusディレクトリでの実行"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            mbox = tmppath / "Test.mbox" / "Messages"
            mbox.mkdir(parents=True)

            src = CORPUS_DIR / "plain_utf8.emlx"
            dst = mbox / "1.emlx"
            dst.write_bytes(src.read_bytes())

            folder = MailFolder(tmppath)
            pattern = SearchPattern("Hello")
            output = tmppath / "output.csv"

            app = MailGrepApp(folder, pattern, output)
            app.run()

            assert output.with_suffix(".csv").exists()
            assert output.with_suffix(".xlsx").exists()

    def test_run_no_matches(self):
        """マッチなしの場合"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            mbox = tmppath / "Test.mbox" / "Messages"
            mbox.mkdir(parents=True)

            src = CORPUS_DIR / "plain_utf8.emlx"
            (mbox / "1.emlx").write_bytes(src.read_bytes())

            folder = MailFolder(tmppath)
            pattern = SearchPattern("NOTFOUNDXYZ")
            output = tmppath / "output.csv"

            app = MailGrepApp(folder, pattern, output)
            app.run()

            assert output.with_suffix(".csv").exists()

    def test_run_multiple_mails(self):
        """複数メールの処理"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            mbox = tmppath / "Test.mbox" / "Messages"
            mbox.mkdir(parents=True)

            for i, name in enumerate(["plain_utf8.emlx", "multi_hit.emlx"]):
                (mbox / f"{i}.emlx").write_bytes((CORPUS_DIR / name).read_bytes())

            folder = MailFolder(tmppath)
            pattern = SearchPattern(".")
            output = tmppath / "output.csv"

            app = MailGrepApp(folder, pattern, output)
            app.run()

            with open(output.with_suffix(".csv"), encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                next(reader)
                rows = list(reader)
                assert len(rows) > 0

    def test_run_empty_folder(self):
        """空フォルダの処理"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)

            folder = MailFolder(tmppath)
            pattern = SearchPattern("test")
            output = tmppath / "output.csv"

            app = MailGrepApp(folder, pattern, output)
            app.run()

            assert output.with_suffix(".csv").exists()
