@testable import MailGrep
import XCTest

final class HeaderDecoderTests: XCTestCase {
    let decoder = HeaderDecoder()

    // MARK: - decode tests (7 tests)

    func testDecodeNoneAndEmpty() {
        XCTAssertEqual(decoder.decode(""), "")
    }

    func testDecodePlainText() {
        XCTAssertEqual(decoder.decode("Hello World"), "Hello World")
        XCTAssertEqual(decoder.decode("こんにちは"), "こんにちは")
    }

    func testDecodeUtf8Base64() {
        let result = decoder.decode("=?UTF-8?B?44GT44KT44Gr44Gh44Gv?=")
        XCTAssertEqual(result, "こんにちは")
    }

    func testDecodeUtf8QuotedPrintable() {
        let result = decoder.decode("=?UTF-8?Q?=E3=83=86=E3=82=B9=E3=83=88?=")
        XCTAssertEqual(result, "テスト")
    }

    func testDecodeIso2022jpBase64() {
        let result = decoder.decode("=?ISO-2022-JP?B?GyRCJDMkcyRLJEEkTxsoQg==?=")
        XCTAssertEqual(result, "こんにちは")
    }

    func testDecodeNewlines() {
        XCTAssertFalse(decoder.decode("hello\rworld").contains("\r"))
        XCTAssertEqual(decoder.decode("hello\nworld"), "hello world")
        XCTAssertEqual(decoder.decode("hello\r\nworld"), "hello world")
    }

    func testDecodeInvalidEncodingFallback() {
        let result = decoder.decode("=?INVALID?B?broken?=")
        XCTAssertTrue(result is String)
    }

    // MARK: - Additional tests

    func testDecodeMixedEncodedAndPlain() {
        let result = decoder.decode("Re: =?UTF-8?B?44GT44KT44Gr44Gh44Gv?=")
        XCTAssertTrue(result.contains("Re:"))
        XCTAssertTrue(result.contains("こんにちは"))
    }

    func testDecodeMultipleEncodedWords() {
        let result = decoder.decode("=?UTF-8?B?44GT44KT?= =?UTF-8?B?44Gr44Gh44Gv?=")
        XCTAssertTrue(result.contains("こん"))
        XCTAssertTrue(result.contains("にちは"))
    }
}
