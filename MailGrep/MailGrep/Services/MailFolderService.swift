import Foundation
import os.log

private let logger = Logger(subsystem: "com.example.MailGrep", category: "MailFolderService")

class MailFolderService {
    private let classifier = MboxClassifier()

    func defaultMailRoot() -> URL {
        let home = FileManager.default.homeDirectoryForCurrentUser
        return home.appendingPathComponent("Library/Mail/V10")
    }

    func collectEmlxFiles(from rootDir: URL? = nil, onlySent: Bool = false) -> [URL] {
        NSLog("📁 [MailFolderService] collectEmlxFiles開始")
        let root = rootDir ?? defaultMailRoot()
        var emlxFiles: [URL] = []

        NSLog("📁 [MailFolderService] メールルートパス: %@", root.path)
        NSLog("📁 [MailFolderService] ルート存在: %d", FileManager.default.fileExists(atPath: root.path) ? 1 : 0)

        guard FileManager.default.fileExists(atPath: root.path) else {
            NSLog("❌ [MailFolderService] ルートディレクトリが存在しません")
            return []
        }

        NSLog("📁 [MailFolderService] mboxディレクトリ検索中...")
        let mboxDirs = findMboxDirectories(in: root)
        NSLog("📁 [MailFolderService] mboxディレクトリ数: %d個", mboxDirs.count)

        var processedCount = 0
        var excludedCount = 0
        for mbox in mboxDirs {
            let mboxName = mbox.lastPathComponent
            if onlySent {
                guard classifier.isSent(mbox) else {
                    NSLog("⏭️ [MailFolderService] スキップ（送信済みでない）: %@", mboxName)
                    excludedCount += 1
                    continue
                }
                NSLog("✅ [MailFolderService] 送信済みフォルダ: %@", mboxName)
            } else {
                if classifier.isExcluded(mbox) {
                    excludedCount += 1
                    continue
                }
                NSLog("✅ [MailFolderService] 処理対象フォルダ: %@", mboxName)
            }

            processedCount += 1
            let emlxInMbox = findEmlxFiles(in: mbox)
            NSLog("📧 [MailFolderService] %@ から %d件のemlxファイル取得", mboxName, emlxInMbox.count)
            emlxFiles.append(contentsOf: emlxInMbox)
        }
        NSLog("📊 [MailFolderService] 処理結果: %d個処理、%d個除外、合計%d件のemlxファイル", processedCount, excludedCount, emlxFiles.count)

        return sortByModificationDate(emlxFiles)
    }

    private func findMboxDirectories(in directory: URL) -> [URL] {
        var result: [URL] = []

        guard let enumerator = FileManager.default.enumerator(
            at: directory,
            includingPropertiesForKeys: [.isDirectoryKey],
            options: [.skipsHiddenFiles]
        ) else {
            NSLog("DEBUG: Failed to create enumerator for %@", directory.path)
            return []
        }

        var enumCount = 0
        while let url = enumerator.nextObject() as? URL {
            enumCount += 1
            if url.pathExtension == "mbox" {
                var isDirectory: ObjCBool = false
                if FileManager.default.fileExists(atPath: url.path, isDirectory: &isDirectory),
                   isDirectory.boolValue {
                    result.append(url)
                }
            }
        }
        NSLog("DEBUG: Enumerated %d items, found %d mbox dirs", enumCount, result.count)

        return result
    }

    private func findEmlxFiles(in mboxDir: URL) -> [URL] {
        var result: [URL] = []
        let mboxName = mboxDir.lastPathComponent

        // mboxディレクトリの直下コンテンツを確認
        do {
            let directContents = try FileManager.default.contentsOfDirectory(at: mboxDir, includingPropertiesForKeys: [.isDirectoryKey], options: [.skipsHiddenFiles])
            NSLog("📂 [findEmlx] %@ 直下のアイテム数: %d", mboxName, directContents.count)
            for item in directContents {
                var isDir: ObjCBool = false
                FileManager.default.fileExists(atPath: item.path, isDirectory: &isDir)
                NSLog("   📄 %@ (%@)", item.lastPathComponent, isDir.boolValue ? "フォルダ" : "ファイル")
            }
        } catch {
            NSLog("❌ [findEmlx] mboxコンテンツ取得失敗: %@", error.localizedDescription)
        }

        guard let enumerator = FileManager.default.enumerator(
            at: mboxDir,
            includingPropertiesForKeys: nil,
            options: [.skipsHiddenFiles]
        ) else {
            NSLog("❌ [findEmlx] enumerator作成失敗: %@", mboxDir.path)
            return []
        }

        var enumCount = 0
        while let url = enumerator.nextObject() as? URL {
            enumCount += 1
            if enumCount <= 3 {
                NSLog("   🔍 列挙中: %@", url.lastPathComponent)
            }
            if url.pathExtension == "emlx" {
                result.append(url)
            }
        }

        NSLog("📊 [findEmlx] %@ : %d個列挙、%d件のemlx発見", mboxName, enumCount, result.count)

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
