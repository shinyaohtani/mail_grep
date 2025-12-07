"""Tests for mbox_classifier.py - MboxClassifier"""
import pytest
import plistlib
import sys
from pathlib import Path
import tempfile

sys.path.insert(0, str(Path(__file__).parent.parent))

from mbox_classifier import MboxClassifier, EXCLUDE_TOKENS, SENT_TOKENS


class TestMboxClassifierNorm:
    """MboxClassifier._norm() のテスト (2個)"""

    def test_norm_basic(self):
        """基本的な正規化"""
        assert MboxClassifier._norm("HELLO") == "hello"
        assert MboxClassifier._norm("hello world!") == "helloworld"

    def test_norm_nfkc(self):
        """NFKC正規化（全角→半角）"""
        assert MboxClassifier._norm("Ｈｅｌｌｏ") == "hello"


class TestMboxClassifierTokens:
    """トークン定数の確認テスト (2個)"""

    def test_exclude_tokens(self):
        """除外トークン"""
        assert "draft" in EXCLUDE_TOKENS
        assert "trash" in EXCLUDE_TOKENS
        assert "junk" in EXCLUDE_TOKENS
        assert "ゴミ箱" in EXCLUDE_TOKENS

    def test_sent_tokens(self):
        """送信済みトークン"""
        assert "sent" in SENT_TOKENS
        assert "送信済み" in SENT_TOKENS


class TestMboxClassifierIsExcluded:
    """MboxClassifier.is_excluded() のテスト (6個)"""

    @pytest.fixture
    def classifier(self):
        return MboxClassifier()

    def test_is_excluded_by_name(self, classifier):
        """フォルダ名で除外判定"""
        with tempfile.TemporaryDirectory() as tmpdir:
            for name in ["Drafts", "Trash", "Junk"]:
                mbox = Path(tmpdir) / f"{name}.mbox"
                mbox.mkdir()
                assert classifier.is_excluded(mbox) is True

    def test_is_excluded_inbox_not_excluded(self, classifier):
        """Inboxは除外されない"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "INBOX.mbox"
            mbox.mkdir()
            assert classifier.is_excluded(mbox) is False

    def test_is_excluded_by_plist_special_attr(self, classifier):
        """Info.plistのMailboxNameで特殊属性による除外判定"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "SomeFolder.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"
            # ホワイトリストのキー（MailboxName）に\\Trashを含める
            plist_data = {"MailboxName": "\\Trash"}
            plist_path.write_bytes(plistlib.dumps(plist_data))
            assert classifier.is_excluded(mbox) is True

    def test_is_excluded_by_plist_name(self, classifier):
        """Info.plistのMailboxNameで除外判定"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "MyFolder.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"
            # ホワイトリストのキー（MailboxName）を使用
            plist_data = {"MailboxName": "下書き"}
            plist_path.write_bytes(plistlib.dumps(plist_data))
            assert classifier.is_excluded(mbox) is True

    def test_is_excluded_japanese_gomi(self, classifier):
        """日本語フォルダ名 ゴミ箱"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "ゴミ箱.mbox"
            mbox.mkdir()
            assert classifier.is_excluded(mbox) is True

    def test_is_excluded_normal_folder(self, classifier):
        """通常フォルダは除外されない"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "Work.mbox"
            mbox.mkdir()
            assert classifier.is_excluded(mbox) is False


class TestMboxClassifierIsSent:
    """MboxClassifier.is_sent() のテスト (5個)"""

    @pytest.fixture
    def classifier(self):
        return MboxClassifier()

    def test_is_sent_by_name(self, classifier):
        """フォルダ名でSent判定"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "Sent.mbox"
            mbox.mkdir()
            assert classifier.is_sent(mbox) is True

    def test_is_sent_japanese(self, classifier):
        """日本語フォルダ名 送信済み"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "送信済み.mbox"
            mbox.mkdir()
            assert classifier.is_sent(mbox) is True

    def test_is_sent_by_plist_special_attr(self, classifier):
        """Info.plistのMailboxNameで\\Sentフラグ"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "MyFolder.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"
            # ホワイトリストのキー（MailboxName）を使用
            plist_data = {"MailboxName": "\\Sent"}
            plist_path.write_bytes(plistlib.dumps(plist_data))
            assert classifier.is_sent(mbox) is True

    def test_is_sent_inbox_false(self, classifier):
        """InboxはSentではない"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "INBOX.mbox"
            mbox.mkdir()
            assert classifier.is_sent(mbox) is False

    def test_is_sent_drafts_false(self, classifier):
        """DraftsはSentではない"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "Drafts.mbox"
            mbox.mkdir()
            assert classifier.is_sent(mbox) is False


class TestMboxClassifierPlistWhitelist:
    """plistホワイトリスト方式の回帰テスト (3個)

    バグ再発防止: ExchangeSyncState等のバイナリデータに含まれる
    偶発的な文字列（"rss", "trash"等）で偽陽性が発生しないことを確認
    """

    @pytest.fixture
    def classifier(self):
        return MboxClassifier()

    def test_inbox_not_excluded_with_rss_in_binary_data(self, classifier):
        """受信トレイがExchangeSyncState内の"rss"で除外されない（回帰テスト）

        実際のバグ: ExchangeSyncStateのbase64データ内に偶然"rss"が含まれ、
        受信トレイが誤って除外されていた
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "受信トレイ.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"
            # 実際のMail.appのplist構造を模倣
            # ExchangeSyncStateにbase64エンコードされた"rss"を含むデータ
            plist_data = {
                "MailboxName": "受信トレイ",
                "MailboxID": "12345",
                "ExchangeSyncState": b"H4sIAAAArssAAAAEAJWYCVTPWf/Hf",  # "rss"を含む
                "FilterEnabled": "NO",
            }
            plist_path.write_bytes(plistlib.dumps(plist_data))
            # 受信トレイは除外されてはいけない
            assert classifier.is_excluded(mbox) is False

    def test_inbox_not_excluded_with_trash_in_sync_data(self, classifier):
        """受信トレイがSyncState内の"trash"で除外されない（回帰テスト）"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "INBOX.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"
            plist_data = {
                "MailboxName": "INBOX",
                "SyncState": "sometrashdata",  # "trash"を含む文字列
                "CachedData": b"junkspamtrashbin",  # 除外トークンを含むバイナリ
            }
            plist_path.write_bytes(plistlib.dumps(plist_data))
            assert classifier.is_excluded(mbox) is False

    def test_only_mailboxname_key_is_used_for_classification(self, classifier):
        """MailboxName以外のキーは分類に使用されない（ホワイトリスト確認）"""
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "MyFolder.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"
            plist_data = {
                "MailboxName": "仕事用",  # 通常のフォルダ名
                "RandomKey": "draft",  # 除外トークンを含むが無視されるべき
                "AnotherKey": "spam",  # 除外トークンを含むが無視されるべき
                "NestedData": {"inner": "trash"},  # ネストされた除外トークン
            }
            plist_path.write_bytes(plistlib.dumps(plist_data))
            # MailboxNameは除外対象ではないので、除外されない
            assert classifier.is_excluded(mbox) is False

    def test_inbox_not_excluded_with_realistic_exchange_sync_state(self, classifier):
        """実際のExchangeSyncState構造を模倣した大きなbase64データでの回帰テスト

        合成データ: gzip圧縮されたXML風データをbase64エンコード
        内部に"rss", "trash", "junk", "spam", "draft", "archive"などを含む
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            mbox = Path(tmpdir) / "受信トレイ.mbox"
            mbox.mkdir()
            plist_path = mbox / "Info.plist"

            # 実際のExchangeSyncStateを模倣する大きな合成データ
            # gzip圧縮されたXML風データで、偶然除外トークンを含む
            large_sync_state = (
                b"H4sIADN2NWkC/+1QwQrCMAy97yuGd12VnaRWhihOPQiKopdRt8xVt06aKPr3toIyv0"
                b"EfgZC8l/ASPrxXpX8Dg6rWg1a3w1pD4fHVQ6crkgTC830+qcsMjOvN4SGmIcaRhUF0"
                b"aRzNtrvRZr3c5jz4VrrZmKB616erPuNFVmQkFgel7QYeNAVuYCGRXC16rBe2WdfGmr"
                b"H+K/Y8+NBOO5JpAc7nFUVmZE5JVWeQ1Lm10uReR6iSrDUgUvqIQpq0UDdIQMtDCZm1"
                b"/s17PGg8Ifrj5/AEMTxyThsDAAA="
            )
            # 追加の大きなバイナリデータ（様々な除外トークンを含む）
            fake_cached_data = (
                b"rss_feed_sync_trash_cleanup_junk_filter_spam_detection_"
                b"draft_autosave_archive_indexer_bin_compactor_" * 50
            )

            plist_data = {
                "MailboxName": "受信トレイ",  # これだけが分類に使用される
                "MailboxID": "AAMkAGE3YjFl...",
                "ExchangeSyncState": large_sync_state,
                "CachedSyncData": fake_cached_data,
                "LastSyncTime": "2024-01-15T10:30:00Z",
                "FilterEnabled": "NO",
                "SyncStateVersion": 3,
                "ItemCount": 1523,
            }
            plist_path.write_bytes(plistlib.dumps(plist_data))

            # 受信トレイは除外されてはいけない
            assert classifier.is_excluded(mbox) is False
            # 受信トレイは送信済みでもない
            assert classifier.is_sent(mbox) is False
