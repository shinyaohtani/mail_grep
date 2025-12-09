@testable import MailGrep
import XCTest

final class EmlxParserTests: XCTestCase {
    let parser = EmlxParser()

    var corpusPath: URL {
        // MailGrep/MailGrepTests -> MailGrep -> mail_grep -> tests/corpus
        let testBundle = Bundle(for: type(of: self))
        let testsDir = URL(fileURLWithPath: #file)
            .deletingLastPathComponent() // MailGrepTests
            .deletingLastPathComponent() // MailGrep
            .deletingLastPathComponent() // mail_grep
            .appendingPathComponent("tests/corpus")
        return testsDir
    }

    // MARK: - Basic parsing tests

    func testParseSimpleEmlx() throws {
        let emlxURL = corpusPath.appendingPathComponent("plain_utf8.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        XCTAssertFalse(parsed.headers.subject.isEmpty)
    }

    func testParseEncodedHeaders() throws {
        let emlxURL = corpusPath.appendingPathComponent("encoded_headers.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        // Encoded headers should be decoded
        XCTAssertFalse(parsed.headers.subject.contains("=?"))
    }

    func testParsePrefixedSize() throws {
        let emlxURL = corpusPath.appendingPathComponent("prefixed_size.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        // Should handle byte size prefix correctly
        XCTAssertNotNil(parsed.headers)
    }

    func testParseHtmlMultipart() throws {
        let emlxURL = corpusPath.appendingPathComponent("html_multipart.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        XCTAssertFalse(parsed.bodyLines.isEmpty)
    }

    func testParseNoDate() throws {
        let emlxURL = corpusPath.appendingPathComponent("no_date.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        // Should handle missing date gracefully
        XCTAssertTrue(parsed.headers.dateStr.isEmpty || parsed.headers.dateValue == nil)
    }

    // MARK: - Header extraction tests

    func testExtractMessageID() throws {
        let emlxURL = corpusPath.appendingPathComponent("plain_utf8.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        // Message-ID should be extracted if present
        XCTAssertNotNil(parsed.headers.messageID)
    }

    func testExtractFromAndTo() throws {
        let emlxURL = corpusPath.appendingPathComponent("plain_utf8.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        XCTAssertNotNil(parsed.headers.from)
        XCTAssertNotNil(parsed.headers.to)
    }

    // MARK: - Body extraction tests

    func testBodyLinesExtraction() throws {
        let emlxURL = corpusPath.appendingPathComponent("multi_hit.emlx")
        guard FileManager.default.fileExists(atPath: emlxURL.path) else {
            XCTFail("Test file not found: \(emlxURL.path)")
            return
        }

        let parsed = try parser.parse(url: emlxURL)
        XCTAssertFalse(parsed.bodyLines.isEmpty)
    }
}
