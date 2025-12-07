import AppKit
import Foundation
import SwiftUI

private let log = CategoryLogger(category: .viewModel)

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
        log.debug("SearchViewModel初期化")
        // デバッグ用：起動時に自動検索
        Task {
            log.debug("起動時自動検索：0.5秒待機")
            try? await Task.sleep(nanoseconds: 500_000_000)
            log.debug("待機完了：search()呼び出し")
            self.search()
        }
    }

    var mailCount: Int {
        Set(results.map { $0.profile.messageID }).count
    }

    func search() {
        log.info("検索開始: パターン='\(pattern)' ignoreCase=\(ignoreCase) onlySent=\(onlySent)")
        guard !pattern.isEmpty else {
            log.warning("パターンが空のため検索中止")
            return
        }

        let searchPattern: SearchPattern
        do {
            searchPattern = try SearchPattern(pattern: pattern, ignoreCase: ignoreCase)
            log.debug("SearchPattern作成成功")
        } catch {
            log.error("SearchPattern作成失敗: \(error.localizedDescription)")
            statusMessage = "無効な正規表現: \(error.localizedDescription)"
            return
        }

        isSearching = true
        searchCompleted = false
        results = []
        progress = 0.0
        statusMessage = "メールを収集中..."
        log.debug("バックグラウンドタスク開始")

        let onlySentCopy = onlySent
        Task.detached { [weak self] in
            await self?.performSearch(searchPattern: searchPattern, onlySent: onlySentCopy)
        }
    }

    private nonisolated func performSearch(searchPattern: SearchPattern, onlySent: Bool) async {
        let searchLog = CategoryLogger(category: .search)
        searchLog.info("performSearch開始: 送信済みのみ=\(onlySent)")

        let localMailFolderService = MailFolderService()
        let localEmlxParser = EmlxParser()

        searchLog.debug("emlxファイル収集中...")
        let emlxFiles = localMailFolderService.collectEmlxFiles(onlySent: onlySent)
        let total = emlxFiles.count
        searchLog.info("emlxファイル総数: \(total)件")

        if total == 0 {
            await MainActor.run { [weak self] in
                self?.statusMessage = "メールが見つかりませんでした"
                self?.isSearching = false
                self?.searchCompleted = true
            }
            searchLog.warning("メールが見つかりませんでした")
            return
        }

        await MainActor.run { [weak self] in
            self?.statusMessage = "\(total)件のメールを検索中..."
        }

        var hitLines: [HitLine] = []
        var mailID = 0
        var lineNumber = 0
        var parseErrors = 0

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
                parseErrors += 1
                // パースエラーは大量に出る可能性があるためdebugレベル
                if parseErrors <= 5 {
                    searchLog.debug("パースエラー[\(parseErrors)]: \(SmartLogger.truncate(emlxURL.path, maxLength: 60))")
                }
            }

            // 進捗更新（100件ごと）
            if index % 100 == 0 || index == total - 1 {
                let currentProgress = Double(index + 1) / Double(total)
                let currentHits = hitLines.count
                await MainActor.run { [weak self] in
                    self?.progress = currentProgress
                    self?.statusMessage = "\(index + 1)/\(total) 検索中... (\(currentHits)件ヒット)"
                }
            }
        }

        // 検索完了
        let finalResults = hitLines
        let finalMailCount = Set(hitLines.map { $0.profile.messageID }).count

        searchLog.info("検索完了: \(finalResults.count)件ヒット（\(finalMailCount)通）、パースエラー: \(parseErrors)件")

        await MainActor.run { [weak self] in
            self?.results = finalResults
            self?.isSearching = false
            self?.searchCompleted = true
            self?.statusMessage = "検索完了: \(finalResults.count)件ヒット（\(finalMailCount)通）"
            self?.progress = 1.0
        }

        Log.finalize()
    }

    func openMailInApp(_ hitLine: HitLine) {
        log.debug("メール開く: \(SmartLogger.truncate(hitLine.profile.subject, maxLength: 30))")
        mailLinkService.openInMailApp(hitLine: hitLine)
    }

    func exportCSV() {
        log.info("CSVエクスポート開始")
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
            log.info("CSV保存成功: \(url.path)")
        } catch {
            log.error("CSV保存エラー: \(error.localizedDescription)")
        }
    }
}
