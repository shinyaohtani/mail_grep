import XCTest

final class ContextMenuTests: XCTestCase {

    var app: XCUIApplication!

    override func setUpWithError() throws {
        continueAfterFailure = false
        app = XCUIApplication()
        app.launch()
    }

    override func tearDownWithError() throws {
        app = nil
    }

    // MARK: - Context Menu Tests

    /// テキストフィールドでテキストを選択し、右クリックでコンテクストメニューを開き、
    /// 「MailGrepで検索...」メニュー項目が存在するかを確認するテスト
    func testContextMenuHasMailGrepSearchOption() throws {
        // 検索パターンのテキストフィールドを探す
        let searchField = app.textFields.firstMatch
        XCTAssertTrue(searchField.waitForExistence(timeout: 5), "検索フィールドが見つかりません")

        // テキストフィールドをクリックしてフォーカス
        searchField.click()

        // テストテキストを入力
        searchField.typeText("テスト検索文字列")

        // テキストを全選択 (Cmd+A)
        searchField.typeKey("a", modifierFlags: .command)

        // 少し待機
        usleep(300_000)  // 0.3秒

        // 右クリックでコンテクストメニューを開く
        searchField.rightClick()

        // コンテクストメニューが表示されるまで待機
        let contextMenu = app.menus.firstMatch
        XCTAssertTrue(contextMenu.waitForExistence(timeout: 3), "コンテクストメニューが表示されません")

        // 「MailGrepで検索...」メニュー項目を探す
        let mailGrepMenuItem = contextMenu.menuItems["MailGrepで検索..."]

        // このテストは現在失敗することを期待（機能未実装のため）
        // 機能実装後はXCTAssertTrueに変更する
        XCTAssertTrue(
            mailGrepMenuItem.exists,
            "「MailGrepで検索...」メニュー項目がコンテクストメニューに存在しません。" +
            "この機能を実装してください。"
        )
    }

    /// コンテクストメニューに存在するすべてのメニュー項目を出力するデバッグ用テスト
    func testDebugPrintContextMenuItems() throws {
        // 検索パターンのテキストフィールドを探す
        let searchField = app.textFields.firstMatch
        XCTAssertTrue(searchField.waitForExistence(timeout: 5), "検索フィールドが見つかりません")

        // テキストフィールドをクリックしてフォーカス
        searchField.click()

        // テストテキストを入力
        searchField.typeText("debug test")

        // テキストを全選択
        searchField.typeKey("a", modifierFlags: .command)
        usleep(300_000)

        // 右クリック
        searchField.rightClick()

        // コンテクストメニューが表示されるまで待機
        let contextMenu = app.menus.firstMatch
        guard contextMenu.waitForExistence(timeout: 3) else {
            XCTFail("コンテクストメニューが表示されません")
            return
        }

        // すべてのメニュー項目を出力
        print("=== コンテクストメニュー項目一覧 ===")
        for menuItem in contextMenu.menuItems.allElementsBoundByIndex {
            let title = menuItem.title
            let identifier = menuItem.identifier
            print("- Title: '\(title)', Identifier: '\(identifier)'")
        }
        print("=== 一覧終了 ===")

        // サブメニュー（サービス等）も確認
        let servicesMenuItem = contextMenu.menuItems["サービス"]
        if servicesMenuItem.exists {
            servicesMenuItem.hover()
            usleep(500_000)

            print("=== サービスサブメニュー項目一覧 ===")
            if let servicesMenu = servicesMenuItem.menus.firstMatch as? XCUIElement,
               servicesMenu.exists {
                for menuItem in servicesMenu.menuItems.allElementsBoundByIndex {
                    print("- Service: '\(menuItem.title)'")
                }
            }
            print("=== サービスサブメニュー一覧終了 ===")
        }

        // このテストは常に成功（デバッグ出力用）
        XCTAssertTrue(true)
    }

    /// Servicesメニュー内に「MailGrepで検索...」があるかを確認するテスト
    func testServicesMenuHasMailGrepSearchOption() throws {
        // 検索パターンのテキストフィールドを探す
        let searchField = app.textFields.firstMatch
        XCTAssertTrue(searchField.waitForExistence(timeout: 5), "検索フィールドが見つかりません")

        // テキストフィールドをクリックしてフォーカス
        searchField.click()

        // テストテキストを入力
        searchField.typeText("サービステスト")

        // テキストを全選択
        searchField.typeKey("a", modifierFlags: .command)
        usleep(300_000)

        // 右クリック
        searchField.rightClick()

        // コンテクストメニューが表示されるまで待機
        let contextMenu = app.menus.firstMatch
        XCTAssertTrue(contextMenu.waitForExistence(timeout: 3), "コンテクストメニューが表示されません")

        // サービスメニューを探す（日本語環境では「サービス」、英語環境では「Services」）
        var servicesMenuItem = contextMenu.menuItems["サービス"]
        if !servicesMenuItem.exists {
            servicesMenuItem = contextMenu.menuItems["Services"]
        }

        guard servicesMenuItem.exists else {
            XCTFail("サービスメニューが見つかりません")
            return
        }

        // サービスメニューにホバーしてサブメニューを開く
        servicesMenuItem.hover()
        usleep(500_000)  // サブメニューが開くまで待機

        // サービスサブメニュー内で「MailGrepで検索...」を探す
        // 注意: XCUITestでサブメニュー項目へのアクセスは制限がある場合がある
        let mailGrepService = app.menuItems["MailGrepで検索..."]

        XCTAssertTrue(
            mailGrepService.exists,
            "「MailGrepで検索...」がサービスメニューに存在しません。" +
            "Info.plistにNSServicesを設定してください。"
        )
    }
}
