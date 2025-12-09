@testable import MailGrep
import XCTest

final class LoggerTests: XCTestCase {
    // MARK: - Log Level Tests

    func testLogLevelComparison() {
        XCTAssertTrue(LogLevel.debug < LogLevel.info)
        XCTAssertTrue(LogLevel.info < LogLevel.warning)
        XCTAssertTrue(LogLevel.warning < LogLevel.error)
    }

    func testLogLevelPrefix() {
        XCTAssertTrue(LogLevel.debug.prefix.contains("DEBUG"))
        XCTAssertTrue(LogLevel.info.prefix.contains("INFO"))
        XCTAssertTrue(LogLevel.warning.prefix.contains("WARN"))
        XCTAssertTrue(LogLevel.error.prefix.contains("ERROR"))
    }

    // MARK: - Truncate Tests

    func testTruncateShortString() {
        let short = "Hello"
        XCTAssertEqual(SmartLogger.truncate(short, maxLength: 50), short)
    }

    func testTruncateLongString() {
        let long = "This is a very long string that should be truncated"
        let truncated = SmartLogger.truncate(long, maxLength: 20)
        XCTAssertEqual(truncated.count, 20)
        XCTAssertTrue(truncated.hasSuffix("..."))
    }

    func testTruncateExactLength() {
        let exact = "12345678901234567890" // 20 chars
        XCTAssertEqual(SmartLogger.truncate(exact, maxLength: 20), exact)
    }

    func testTruncateDefaultLength() {
        let long = String(repeating: "a", count: 100)
        let truncated = SmartLogger.truncate(long) // default maxLength: 50
        XCTAssertEqual(truncated.count, 50)
        XCTAssertTrue(truncated.hasSuffix("..."))
    }

    // MARK: - Elapsed Time Tests

    func testElapsedTimeFormat() {
        let elapsed = Log.elapsedTimeString()
        // Format: h:mm:ss
        let pattern = #"^\d+:\d{2}:\d{2}$"#
        XCTAssertNotNil(elapsed.range(of: pattern, options: .regularExpression))
    }

    func testElapsedSeconds() {
        let seconds = Log.elapsedSeconds
        XCTAssertGreaterThanOrEqual(seconds, 0)
    }

    // MARK: - CategoryLogger Tests

    func testCategoryLoggerCreation() {
        let viewModelLog = CategoryLogger(category: .viewModel)
        let serviceLog = CategoryLogger(category: .service)
        let classifierLog = CategoryLogger(category: .classifier)

        // Just verify they can be created without crashing
        XCTAssertNotNil(viewModelLog)
        XCTAssertNotNil(serviceLog)
        XCTAssertNotNil(classifierLog)
    }

    // MARK: - LogCategory Tests

    func testLogCategoryRawValues() {
        XCTAssertEqual(LogCategory.app.rawValue, "App")
        XCTAssertEqual(LogCategory.viewModel.rawValue, "ViewModel")
        XCTAssertEqual(LogCategory.parser.rawValue, "Parser")
        XCTAssertEqual(LogCategory.service.rawValue, "Service")
        XCTAssertEqual(LogCategory.classifier.rawValue, "Classifier")
        XCTAssertEqual(LogCategory.search.rawValue, "Search")
    }

    func testLogCategorySubsystem() {
        XCTAssertEqual(LogCategory.app.subsystem, "com.mailgrep.app")
    }

    // MARK: - Global Logger Tests

    func testGlobalLoggerExists() {
        XCTAssertNotNil(Log)
        XCTAssertTrue(Log === SmartLogger.shared)
    }

    // MARK: - Debug Mode Tests

    func testDebugModeDefault() {
        // デフォルトではデバッグモードはオフ（環境変数が設定されていない場合）
        // 環境変数の状態に依存するため、この状態を明示的にテストするのは難しい
        // ただし、プロパティがアクセス可能であることを確認
        _ = Log.isDebugMode
    }

    // MARK: - Log Level Setting Tests

    func testSetLevel() {
        Log.setLevel(.warning)
        XCTAssertEqual(Log.effectiveLevel, Log.isDebugMode ? .debug : .warning)

        // Reset to default
        Log.setLevel(.info)
    }
}
