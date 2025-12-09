import XCTest

final class SearchFunctionTests: XCTestCase {
    var app: XCUIApplication!

    override func setUpWithError() throws {
        continueAfterFailure = false
        app = XCUIApplication()
        app.launch()
    }

    override func tearDownWithError() throws {
        app = nil
    }

    // MARK: - 基本検索機能テスト

    /// 検索ボタンをクリックして検索が実行されるかテスト
    func testSearchButtonTriggersSearch() throws {
        // 検索パターンのテキストフィールドを探す
        let searchField = app.textFields["searchPatternTextField"]
        XCTAssertTrue(searchField.waitForExistence(timeout: 5), "検索フィールドが見つかりません")

        // 現在のテキストをクリアして新しいパターンを入力
        searchField.click()
        searchField.typeKey("a", modifierFlags: .command) // 全選択
        searchField.typeText("test") // テスト用パターン

        // 検索ボタンを探してクリック
        let searchButton = app.buttons["検索"]
        XCTAssertTrue(searchButton.waitForExistence(timeout: 3), "検索ボタンが見つかりません")
        searchButton.click()

        // 検索が開始されたことを確認（プログレス表示または結果表示を待つ）
        // 少し待機して検索が実行される時間を与える
        sleep(3)

        // 検索結果または「検索完了」「検索結果なし」のステータスが表示されるか確認
        let resultText = app.staticTexts.matching(NSPredicate(format: "label CONTAINS '結果' OR label CONTAINS '検索中' OR label CONTAINS '完了'")).firstMatch
        XCTAssertTrue(resultText.waitForExistence(timeout: 10), "検索結果またはステータスが表示されません。検索が正しく実行されていない可能性があります。")
    }

    /// Enterキーで検索が実行されるかテスト
    func testEnterKeyTriggersSearch() throws {
        // 検索パターンのテキストフィールドを探す
        let searchField = app.textFields["searchPatternTextField"]
        XCTAssertTrue(searchField.waitForExistence(timeout: 5), "検索フィールドが見つかりません")

        // 現在のテキストをクリアして新しいパターンを入力
        searchField.click()
        searchField.typeKey("a", modifierFlags: .command) // 全選択
        searchField.typeText("hello") // テスト用パターン

        // Enterキーを押して検索
        searchField.typeKey(.return, modifierFlags: [])

        // 検索が開始されたことを確認
        sleep(3)

        // 検索結果またはステータスが表示されるか確認
        let resultText = app.staticTexts.matching(NSPredicate(format: "label CONTAINS '結果' OR label CONTAINS '検索中' OR label CONTAINS '完了'")).firstMatch
        XCTAssertTrue(resultText.waitForExistence(timeout: 10), "Enterキーでの検索が正しく実行されていない可能性があります。")
    }

    /// サービス通知からの検索が実行されるかテスト（シミュレーション）
    /// 注：このテストはNotificationCenterを直接テストできないため、アプリが起動時に自動検索を行うことを確認
    func testAppStartsWithAutoSearch() throws {
        // アプリは起動時にデフォルトパターン "test" で自動検索を開始する
        // 検索が開始されることを確認

        // 検索結果またはステータスが表示されるまで待つ
        let resultText = app.staticTexts.matching(NSPredicate(format: "label CONTAINS '結果' OR label CONTAINS '検索中' OR label CONTAINS '完了' OR label CONTAINS 'メール'")).firstMatch

        // 起動時自動検索が実行されるため、10秒以内に何らかのステータスが表示されるはず
        XCTAssertTrue(resultText.waitForExistence(timeout: 15), "起動時の自動検索が正しく実行されていない可能性があります。")
    }
}
