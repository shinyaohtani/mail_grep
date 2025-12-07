import XCTest
@testable import MailGrep

final class MboxClassifierTests: XCTestCase {

    let classifier = MboxClassifier()

    // MARK: - isExcluded tests

    func testIsExcludedByName() {
        let trashURL = URL(fileURLWithPath: "/tmp/test/Trash.mbox")
        // Note: This tests the name-based fallback since no Info.plist exists
        XCTAssertTrue(classifier.isExcluded(trashURL))
    }

    func testIsExcludedDrafts() {
        let draftsURL = URL(fileURLWithPath: "/tmp/test/Drafts.mbox")
        XCTAssertTrue(classifier.isExcluded(draftsURL))
    }

    func testIsExcludedJunk() {
        let junkURL = URL(fileURLWithPath: "/tmp/test/Junk.mbox")
        XCTAssertTrue(classifier.isExcluded(junkURL))
    }

    func testIsExcludedSpam() {
        let spamURL = URL(fileURLWithPath: "/tmp/test/Spam.mbox")
        XCTAssertTrue(classifier.isExcluded(spamURL))
    }

    func testIsExcludedArchive() {
        let archiveURL = URL(fileURLWithPath: "/tmp/test/Archive.mbox")
        XCTAssertTrue(classifier.isExcluded(archiveURL))
    }

    func testIsExcludedJapaneseGomi() {
        let gomiURL = URL(fileURLWithPath: "/tmp/test/ゴミ箱.mbox")
        XCTAssertTrue(classifier.isExcluded(gomiURL))
    }

    func testIsExcludedJapaneseMeiwaku() {
        let meiwakuURL = URL(fileURLWithPath: "/tmp/test/迷惑メール.mbox")
        XCTAssertTrue(classifier.isExcluded(meiwakuURL))
    }

    func testInboxNotExcluded() {
        let inboxURL = URL(fileURLWithPath: "/tmp/test/INBOX.mbox")
        XCTAssertFalse(classifier.isExcluded(inboxURL))
    }

    func testNormalFolderNotExcluded() {
        let normalURL = URL(fileURLWithPath: "/tmp/test/Work.mbox")
        XCTAssertFalse(classifier.isExcluded(normalURL))
    }

    // MARK: - isSent tests

    func testIsSentByName() {
        let sentURL = URL(fileURLWithPath: "/tmp/test/Sent.mbox")
        XCTAssertTrue(classifier.isSent(sentURL))
    }

    func testIsSentJapanese() {
        let sentURL = URL(fileURLWithPath: "/tmp/test/送信済み.mbox")
        XCTAssertTrue(classifier.isSent(sentURL))
    }

    func testIsSentGerman() {
        let sentURL = URL(fileURLWithPath: "/tmp/test/Gesendet.mbox")
        XCTAssertTrue(classifier.isSent(sentURL))
    }

    func testInboxNotSent() {
        let inboxURL = URL(fileURLWithPath: "/tmp/test/INBOX.mbox")
        XCTAssertFalse(classifier.isSent(inboxURL))
    }

    func testDraftsNotSent() {
        let draftsURL = URL(fileURLWithPath: "/tmp/test/Drafts.mbox")
        XCTAssertFalse(classifier.isSent(draftsURL))
    }

    // MARK: - Plist Whitelist Regression Tests (バグ再発防止テスト)

    /// 受信トレイがExchangeSyncState内の"rss"で除外されない（回帰テスト）
    /// 実際のバグ: ExchangeSyncStateのbase64データ内に偶然"rss"が含まれ、
    /// 受信トレイが誤って除外されていた
    func testInboxNotExcludedWithRssInBinaryData() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
        let mboxDir = tempDir.appendingPathComponent("受信トレイ.mbox")
        try FileManager.default.createDirectory(at: mboxDir, withIntermediateDirectories: true)

        // 実際のMail.appのplist構造を模倣
        // ExchangeSyncStateにbase64エンコードされた"rss"を含むデータ
        let plistData: [String: Any] = [
            "MailboxName": "受信トレイ",
            "MailboxID": "12345",
            "ExchangeSyncState": "H4sIAAAArssAAAAEAJWYCVTPWf/Hf",  // "rss"を含む
            "FilterEnabled": "NO"
        ]
        let plistPath = mboxDir.appendingPathComponent("Info.plist")
        let data = try PropertyListSerialization.data(fromPropertyList: plistData, format: .xml, options: 0)
        try data.write(to: plistPath)

        // 受信トレイは除外されてはいけない
        XCTAssertFalse(classifier.isExcluded(mboxDir))

        // クリーンアップ
        try? FileManager.default.removeItem(at: tempDir)
    }

    /// 受信トレイがSyncState内の"trash"で除外されない（回帰テスト）
    func testInboxNotExcludedWithTrashInSyncData() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
        let mboxDir = tempDir.appendingPathComponent("INBOX.mbox")
        try FileManager.default.createDirectory(at: mboxDir, withIntermediateDirectories: true)

        let plistData: [String: Any] = [
            "MailboxName": "INBOX",
            "SyncState": "sometrashdata",  // "trash"を含む文字列
            "CachedData": "junkspamtrashbin"  // 除外トークンを含む
        ]
        let plistPath = mboxDir.appendingPathComponent("Info.plist")
        let data = try PropertyListSerialization.data(fromPropertyList: plistData, format: .xml, options: 0)
        try data.write(to: plistPath)

        XCTAssertFalse(classifier.isExcluded(mboxDir))

        try? FileManager.default.removeItem(at: tempDir)
    }

    /// MailboxName以外のキーは分類に使用されない（ホワイトリスト確認）
    func testOnlyMailboxNameKeyIsUsedForClassification() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
        let mboxDir = tempDir.appendingPathComponent("MyFolder.mbox")
        try FileManager.default.createDirectory(at: mboxDir, withIntermediateDirectories: true)

        let plistData: [String: Any] = [
            "MailboxName": "仕事用",  // 通常のフォルダ名
            "RandomKey": "draft",  // 除外トークンを含むが無視されるべき
            "AnotherKey": "spam",  // 除外トークンを含むが無視されるべき
            "NestedData": ["inner": "trash"]  // ネストされた除外トークン
        ]
        let plistPath = mboxDir.appendingPathComponent("Info.plist")
        let data = try PropertyListSerialization.data(fromPropertyList: plistData, format: .xml, options: 0)
        try data.write(to: plistPath)

        // MailboxNameは除外対象ではないので、除外されない
        XCTAssertFalse(classifier.isExcluded(mboxDir))

        try? FileManager.default.removeItem(at: tempDir)
    }

    /// 実際のExchangeSyncState構造を模倣した大きなbase64データでの回帰テスト
    /// 合成データ: gzip圧縮されたXML風データをbase64エンコード
    /// 内部に"rss", "trash", "junk", "spam", "draft", "archive"などを含む
    func testInboxNotExcludedWithRealisticExchangeSyncState() throws {
        let tempDir = FileManager.default.temporaryDirectory
            .appendingPathComponent(UUID().uuidString)
        let mboxDir = tempDir.appendingPathComponent("受信トレイ.mbox")
        try FileManager.default.createDirectory(at: mboxDir, withIntermediateDirectories: true)

        // 実際のExchangeSyncStateを模倣する大きな合成データ
        // gzip圧縮されたXML風データで、偶然除外トークンを含む
        let largeSyncState = "H4sIADN2NWkC/+1QwQrCMAy97yuGd12VnaRWhihOPQiKopdRt8xVt06aKPr3toIyv0" +
            "EfgZC8l/ASPrxXpX8Dg6rWg1a3w1pD4fHVQ6crkgTC830+qcsMjOvN4SGmIcaRhUF0" +
            "aRzNtrvRZr3c5jz4VrrZmKB616erPuNFVmQkFgel7QYeNAVuYCGRXC16rBe2WdfGmr" +
            "H+K/Y8+NBOO5JpAc7nFUVmZE5JVWeQ1Lm10uReR6iSrDUgUvqIQpq0UDdIQMtDCZm1" +
            "/s17PGg8Ifrj5/AEMTxyThsDAAA="

        // 追加の大きな文字列データ（様々な除外トークンを含む）
        let fakeCachedData = String(repeating: "rss_feed_sync_trash_cleanup_junk_filter_spam_detection_draft_autosave_archive_indexer_bin_compactor_", count: 50)

        let plistData: [String: Any] = [
            "MailboxName": "受信トレイ",  // これだけが分類に使用される
            "MailboxID": "AAMkAGE3YjFl...",
            "ExchangeSyncState": largeSyncState,
            "CachedSyncData": fakeCachedData,
            "LastSyncTime": "2024-01-15T10:30:00Z",
            "FilterEnabled": "NO",
            "SyncStateVersion": 3,
            "ItemCount": 1523
        ]
        let plistPath = mboxDir.appendingPathComponent("Info.plist")
        let data = try PropertyListSerialization.data(fromPropertyList: plistData, format: .xml, options: 0)
        try data.write(to: plistPath)

        // 受信トレイは除外されてはいけない
        XCTAssertFalse(classifier.isExcluded(mboxDir))
        // 受信トレイは送信済みでもない
        XCTAssertFalse(classifier.isSent(mboxDir))

        try? FileManager.default.removeItem(at: tempDir)
    }
}
