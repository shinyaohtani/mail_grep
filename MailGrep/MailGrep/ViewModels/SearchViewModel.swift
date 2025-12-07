import AppKit
import Foundation
import SwiftUI

@MainActor
class SearchViewModel: ObservableObject {
    @Published var pattern: String = "test"
    @Published var ignoreCase: Bool = true
    @Published var onlySent: Bool = false
    @Published var results: [HitLine] = []
    @Published var isSearching: Bool = false
    @Published var progress: Double = 0.0
    @Published var statusMessage: String = ""
    @Published var searchCompleted: Bool = false

    private let mailFolderService = MailFolderService()
    private let mailLinkService = MailLinkService()
    private let emlxParser = EmlxParser()

    init() {
        NSLog("📱 [ViewModel] init() が呼ばれました")
        // デバッグ用：起動時に自動検索
        Task {
            NSLog("📱 [ViewModel] Task開始：0.5秒待機します")
            try? await Task.sleep(nanoseconds: 500_000_000) // 0.5秒待機
            NSLog("📱 [ViewModel] 待機完了：search()を呼び出します")
            self.search()
        }
    }

    var mailCount: Int {
        Set(results.map { $0.profile.messageID }).count
    }

    func search() {
        NSLog("🔍 [ViewModel] search() が呼ばれました。パターン: %@", pattern)
        guard !pattern.isEmpty else {
            NSLog("⚠️ [ViewModel] パターンが空のため検索を中止")
            return
        }

        let searchPattern: SearchPattern
        do {
            searchPattern = try SearchPattern(pattern: pattern, ignoreCase: ignoreCase)
            NSLog("✅ [ViewModel] SearchPattern作成成功")
        } catch {
            NSLog("❌ [ViewModel] SearchPattern作成失敗: %@", error.localizedDescription)
            statusMessage = "無効な正規表現: \(error.localizedDescription)"
            return
        }

        isSearching = true
        searchCompleted = false
        results = []
        progress = 0.0
        statusMessage = "メールを収集中..."
        NSLog("🚀 [ViewModel] バックグラウンドタスクを開始します")

        let onlySentCopy = onlySent
        Task.detached { [weak self] in
            NSLog("🏃 [ViewModel] detachedタスク内：performSearchを呼び出します")
            await self?.performSearch(searchPattern: searchPattern, onlySent: onlySentCopy)
        }
    }

    private nonisolated func performSearch(searchPattern: SearchPattern, onlySent: Bool) async {
        NSLog("📂 [検索] performSearch開始。送信済みのみ=%d", onlySent ? 1 : 0)
        let localMailFolderService = MailFolderService()
        let localEmlxParser = EmlxParser()

        NSLog("📂 [検索] collectEmlxFiles呼び出し中...")
        let emlxFiles = localMailFolderService.collectEmlxFiles(onlySent: onlySent)
        let total = emlxFiles.count
        NSLog("📂 [検索] emlxファイル総数: %d件", total)

        if total == 0 {
            await MainActor.run { [weak self] in
                self?.statusMessage = "メールが見つかりませんでした"
                self?.isSearching = false
                self?.searchCompleted = true
            }
            return
        }

        await MainActor.run { [weak self] in
            self?.statusMessage = "\(total)件のメールを検索中..."
        }

        var hitLines: [HitLine] = []
        var mailID = 0
        var lineNumber = 0

        for (index, emlxURL) in emlxFiles.enumerated() {
            do {
                let parsed = try localEmlxParser.parse(url: emlxURL)
                var foundInThisMail = false
                var thisMailID = 0

                let profile = MailProfile(
                    messageID: parsed.headers.messageID,
                    dateStr: parsed.headers.dateStr,
                    dateValue: parsed.headers.dateValue,
                    link: "",
                    subject: parsed.headers.subject,
                    fromAddr: parsed.headers.from,
                    toAddr: parsed.headers.to,
                    emlxPath: emlxURL
                )

                for headerLine in parsed.headers.headerLines {
                    if searchPattern.matches(headerLine) {
                        if !foundInThisMail {
                            mailID += 1
                            thisMailID = mailID
                            foundInThisMail = true
                        }
                        lineNumber += 1
                        let hitLine = HitLine(
                            mailID: thisMailID,
                            profile: profile,
                            matchedLine: headerLine,
                            lineNumber: lineNumber
                        )
                        hitLines.append(hitLine)
                    }
                }

                for (line, type) in parsed.bodyLines {
                    if type == "text/plain" || type == "text/html_textonly" {
                        if searchPattern.matches(line) {
                            if !foundInThisMail {
                                mailID += 1
                                thisMailID = mailID
                                foundInThisMail = true
                            }
                            lineNumber += 1
                            let hitLine = HitLine(
                                mailID: thisMailID,
                                profile: profile,
                                matchedLine: line,
                                lineNumber: lineNumber
                            )
                            hitLines.append(hitLine)
                        }
                    }
                }
            } catch {
                // Skip files that can't be parsed
            }

            if index % 100 == 0 || index == total - 1 {
                let currentProgress = Double(index + 1) / Double(total)
                let currentHits = hitLines.count
                await MainActor.run { [weak self] in
                    self?.progress = currentProgress
                    self?.statusMessage = "\(index + 1)/\(total) 検索中... (\(currentHits)件ヒット)"
                }
            }
        }

        let finalResults = hitLines
        await MainActor.run { [weak self] in
            self?.results = finalResults
            self?.isSearching = false
            self?.searchCompleted = true
            self?.statusMessage = "検索完了"
            self?.progress = 1.0
        }
    }

    func openMailInApp(_ hitLine: HitLine) {
        mailLinkService.openInMailApp(hitLine: hitLine)
    }

    func exportCSV() {
        let panel = NSSavePanel()
        panel.allowedContentTypes = [.commaSeparatedText]
        panel.nameFieldStringValue = "search_results.csv"

        panel.begin { response in
            if response == .OK, let url = panel.url {
                self.saveCSV(to: url)
            }
        }
    }

    private func saveCSV(to url: URL) {
        var csv = "\u{FEFF}"  // BOM for UTF-8
        csv += "No,日付,件名,From,To,マッチ行\n"

        for hitLine in results {
            csv += hitLine.csvRow() + "\n"
        }

        do {
            try csv.write(to: url, atomically: true, encoding: .utf8)
        } catch {
            print("CSV保存エラー: \(error)")
        }
    }
}
