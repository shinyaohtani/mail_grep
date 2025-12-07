import Foundation

private let log = CategoryLogger(category: .service)

class MailFolderService {
    private let classifier = MboxClassifier()

    func defaultMailRoot() -> URL {
        let home = FileManager.default.homeDirectoryForCurrentUser
        return home.appendingPathComponent("Library/Mail/V10")
    }

    func collectEmlxFiles(from rootDir: URL? = nil, onlySent: Bool = false) -> [URL] {
        log.info("collectEmlxFiles開始")
        let root = rootDir ?? defaultMailRoot()
        var emlxFiles: [URL] = []

        log.debug("メールルートパス: \(root.path)")

        guard FileManager.default.fileExists(atPath: root.path) else {
            log.error("ルートディレクトリが存在しません: \(root.path)")
            return []
        }

        log.debug("mboxディレクトリ検索中...")
        let mboxDirs = findMboxDirectories(in: root)
        log.info("mboxディレクトリ数: \(mboxDirs.count)個")

        var processedCount = 0
        var excludedCount = 0
        for mbox in mboxDirs {
            let mboxName = mbox.lastPathComponent
            if onlySent {
                guard classifier.isSent(mbox) else {
                    excludedCount += 1
                    continue
                }
                log.debug("送信済みフォルダ: \(mboxName)")
            } else {
                if classifier.isExcluded(mbox) {
                    excludedCount += 1
                    continue
                }
            }

            processedCount += 1
            let emlxInMbox = findEmlxFiles(in: mbox)
            if emlxInMbox.count > 0 {
                log.debug("\(mboxName): \(emlxInMbox.count)件")
            }
            emlxFiles.append(contentsOf: emlxInMbox)
        }
        log.info("処理結果: \(processedCount)個処理、\(excludedCount)個除外、合計\(emlxFiles.count)件のemlx")

        return sortByModificationDate(emlxFiles)
    }

    private func findMboxDirectories(in directory: URL) -> [URL] {
        var result: [URL] = []

        guard let enumerator = FileManager.default.enumerator(
            at: directory,
            includingPropertiesForKeys: [.isDirectoryKey],
            options: [.skipsHiddenFiles]
        ) else {
            log.warning("enumerator作成失敗: \(directory.path)")
            return []
        }

        while let url = enumerator.nextObject() as? URL {
            if url.pathExtension == "mbox" {
                var isDirectory: ObjCBool = false
                if FileManager.default.fileExists(atPath: url.path, isDirectory: &isDirectory),
                   isDirectory.boolValue {
                    result.append(url)
                }
            }
        }

        return result
    }

    private func findEmlxFiles(in mboxDir: URL) -> [URL] {
        var result: [URL] = []

        guard let enumerator = FileManager.default.enumerator(
            at: mboxDir,
            includingPropertiesForKeys: nil,
            options: [.skipsHiddenFiles]
        ) else {
            log.warning("enumerator作成失敗: \(mboxDir.path)")
            return []
        }

        while let url = enumerator.nextObject() as? URL {
            if url.pathExtension == "emlx" {
                result.append(url)
            }
        }

        return result
    }

    private func sortByModificationDate(_ files: [URL]) -> [URL] {
        return files.sorted { url1, url2 in
            let date1 = (try? url1.resourceValues(forKeys: [.contentModificationDateKey]).contentModificationDate) ?? Date.distantPast
            let date2 = (try? url2.resourceValues(forKeys: [.contentModificationDateKey]).contentModificationDate) ?? Date.distantPast
            return date1 > date2
        }
    }
}
