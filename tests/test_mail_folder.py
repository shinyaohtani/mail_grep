"""Tests for mail_folder.py - MailFolder"""
import pytest
import plistlib
import sys
from pathlib import Path
import tempfile
import time

sys.path.insert(0, str(Path(__file__).parent.parent))

from mail_folder import MailFolder


def create_mbox_structure(tmpdir: Path, name: str, emlx_count: int = 1, plist_data: dict = None):
    """テスト用mboxディレクトリ構造を作成"""
    mbox = tmpdir / f"{name}.mbox"
    mbox.mkdir(parents=True)

    if plist_data:
        (mbox / "Info.plist").write_bytes(plistlib.dumps(plist_data))

    messages_dir = mbox / "Messages"
    messages_dir.mkdir()

    for i in range(emlx_count):
        emlx = messages_dir / f"message_{i}.emlx"
        emlx.write_text(f"From: test{i}@example.com\nSubject: Test {i}\n\nBody {i}")
        time.sleep(0.01)

    return mbox


class TestMailFolderBasic:
    """MailFolder 基本機能のテスト (4個)"""

    def test_create_mail_folder(self):
        """MailFolderの生成"""
        with tempfile.TemporaryDirectory() as tmpdir:
            folder = MailFolder(Path(tmpdir))
            assert folder._root_dir == Path(tmpdir)
            assert folder._only_sent is False

    def test_mail_paths_empty_dir(self):
        """空ディレクトリでは空リスト"""
        with tempfile.TemporaryDirectory() as tmpdir:
            folder = MailFolder(Path(tmpdir))
            paths = folder.mail_paths()
            assert paths == []

    def test_mail_paths_finds_emlx(self):
        """emlxファイルを発見"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=3)

            folder = MailFolder(tmppath)
            paths = folder.mail_paths()
            assert len(paths) == 3
            assert all(p.suffix == ".emlx" for p in paths)

    def test_mail_paths_sorted_by_mtime(self):
        """更新時刻の降順でソート"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=3)

            folder = MailFolder(tmppath)
            paths = folder.mail_paths()
            mtimes = [p.stat().st_mtime for p in paths]
            assert mtimes == sorted(mtimes, reverse=True)


class TestMailFolderExclusion:
    """MailFolder 除外機能のテスト (3個)"""

    def test_mail_paths_excludes_special_folders(self):
        """特殊フォルダは除外"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=2)
            create_mbox_structure(tmppath, "Trash", emlx_count=3)
            create_mbox_structure(tmppath, "Junk", emlx_count=3)

            folder = MailFolder(tmppath)
            paths = folder.mail_paths()
            assert len(paths) == 2

    def test_mail_paths_excludes_by_plist(self):
        """Info.plistの属性で除外"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=2)
            create_mbox_structure(
                tmppath, "MyTrash", emlx_count=3,
                plist_data={"SpecialUseFlags": ["\\Trash"]}
            )

            folder = MailFolder(tmppath)
            paths = folder.mail_paths()
            assert len(paths) == 2

    def test_mail_paths_includes_normal_folder(self):
        """通常フォルダは含まれる"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Work", emlx_count=2)
            create_mbox_structure(tmppath, "Personal", emlx_count=3)

            folder = MailFolder(tmppath)
            paths = folder.mail_paths()
            assert len(paths) == 5


class TestMailFolderOnlySent:
    """MailFolder only_sent モードのテスト (3個)"""

    def test_only_sent_filters_sent(self):
        """only_sent=TrueでSentのみ取得"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=2)
            create_mbox_structure(tmppath, "Sent", emlx_count=3)

            folder = MailFolder(tmppath, only_sent=True)
            paths = folder.mail_paths()
            assert len(paths) == 3

    def test_only_sent_by_plist(self):
        """Info.plistの\\Sentフラグで判定"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=2)
            create_mbox_structure(
                tmppath, "MySent", emlx_count=3,
                plist_data={"SpecialUseFlags": ["\\Sent"]}
            )

            folder = MailFolder(tmppath, only_sent=True)
            paths = folder.mail_paths()
            assert len(paths) == 3

    def test_only_sent_empty_when_no_sent(self):
        """Sentフォルダがない場合は空"""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)
            create_mbox_structure(tmppath, "Inbox", emlx_count=2)

            folder = MailFolder(tmppath, only_sent=True)
            paths = folder.mail_paths()
            assert len(paths) == 0
