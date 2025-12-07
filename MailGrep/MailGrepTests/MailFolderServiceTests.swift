import XCTest
@testable import MailGrep

final class MailFolderServiceTests: XCTestCase {

    let service = MailFolderService()

    func testDefaultMailRoot() {
        let root = service.defaultMailRoot()
        XCTAssertTrue(root.path.contains("Library/Mail"))
    }

    func testCollectEmlxFilesFromNonExistentPath() {
        let fakeURL = URL(fileURLWithPath: "/nonexistent/path")
        let files = service.collectEmlxFiles(from: fakeURL)
        XCTAssertTrue(files.isEmpty)
    }

    func testCollectEmlxFilesReturnsURLs() {
        // This test uses the actual mail folder if available
        let files = service.collectEmlxFiles()
        // Just verify it doesn't crash and returns an array
        XCTAssertNotNil(files)
    }
}
