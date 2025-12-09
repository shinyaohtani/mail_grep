import AppKit
import Foundation
import SwiftUI

private let log = CategoryLogger(category: .viewModel)

@MainActor
class SearchViewModel: ObservableObject {
    @Published var pattern: String = ""
    @Published var ignoreCase: Bool = true
    @Published var onlySent: Bool = false
    @Published var results: [HitLine] = []
    @Published var isSearching: Bool = false
    @Published var progress: Double = 0.0
    @Published var statusMessage: String = ""
    @Published var searchCompleted: Bool = false
    @Published var searchHistory: [String] = []

    private let mailFolderService = MailFolderService()
    private let mailLinkService = MailLinkService()
    private let emlxParser = EmlxParser()

    /// 現在の検索タスク（キャンセル用）
    private var currentSearchTask: Task<Void, Never>?

    // UserDefaults keys
    private static let currentKeywordKey = "currentSearchKeyword"
    private static let searchHistoryKey = "searchHistory"
    private static let normalTerminationKey = "normalTermination"
    private static let maxHistoryCount = 20

    init() {
        log.debug("SearchViewModel初期化")
        loadState()
    }

    /// 状態をUserDefaultsから読み込む
    private func loadState() {
        let defaults = UserDefaults.standard

        // 通常終了フラグをチェック
        let wasNormalTermination = defaults.bool(forKey: Self.normalTerminationKey)

        if wasNormalTermination {
            // 通常終了後はキーワードをクリア（プレースホルダー表示）
            pattern = ""
            log.debug("通常終了後：キーワードをクリア")
        } else {
            // 異常終了後は前回のキーワードを復元
            if let savedKeyword = defaults.string(forKey: Self.currentKeywordKey), !savedKeyword.isEmpty {
                pattern = savedKeyword
                log.debug("異常終了後：キーワード復元 '\(savedKeyword)'")
            }
        }

        // 検索履歴を読み込む
        if let history = defaults.stringArray(forKey: Self.searchHistoryKey) {
            searchHistory = history
            log.debug("検索履歴読み込み: \(history.count)件")
        }

        // 通常終了フラグをリセット（次回起動時に異常終了とみなす）
        defaults.set(false, forKey: Self.normalTerminationKey)
    }

    /// 現在のキーワードを保存（定期的に呼ばれる）
    func saveCurrentKeyword() {
        UserDefaults.standard.set(pattern, forKey: Self.currentKeywordKey)
    }

    /// 通常終了時の処理
    func prepareForNormalTermination() {
        let defaults = UserDefaults.standard
        defaults.set(true, forKey: Self.normalTerminationKey)
        defaults.removeObject(forKey: Self.currentKeywordKey)
        log.debug("通常終了準備：キーワードをクリア")
    }

    /// 検索履歴に追加
    private func addToHistory(_ keyword: String) {
        guard !keyword.isEmpty else { return }

        // 既に存在する場合は削除して先頭に追加
        searchHistory.removeAll { $0 == keyword }
        searchHistory.insert(keyword, at: 0)

        // 最大件数を超えた分を削除
        if searchHistory.count > Self.maxHistoryCount {
            searchHistory = Array(searchHistory.prefix(Self.maxHistoryCount))
        }

        // 保存
        UserDefaults.standard.set(searchHistory, forKey: Self.searchHistoryKey)
        log.debug("検索履歴追加: '\(keyword)' (計\(searchHistory.count)件)")
    }

    /// 検索履歴をクリア
    func clearHistory() {
        searchHistory = []
        UserDefaults.standard.removeObject(forKey: Self.searchHistoryKey)
        log.debug("検索履歴クリア")
    }

    /// 履歴から検索キーワードを選択（テキストボックスに入力するだけ）
    func selectFromHistory(_ keyword: String) {
        pattern = keyword
    }

    var mailCount: Int {
        Set(results.map(\.profile.messageID)).count
    }

    func search() {
        log.info("検索開始: パターン='\(pattern)' ignoreCase=\(ignoreCase) onlySent=\(onlySent)")

        // 権限チェック
        if !PermissionChecker.shared.hasMailAccess() {
            log.warning("権限がないため検索中止")
            PermissionChecker.shared.showPermissionAlertIfNeeded()
            return
        }

        guard !pattern.isEmpty else {
            log.warning("パターンが空のため検索中止")
            statusMessage = "エラー: パターンが空です"
            return
        }

        // 既存の検索をキャンセル
        if let existingTask = currentSearchTask {
            log.info("既存の検索をキャンセル")
            existingTask.cancel()
            currentSearchTask = nil
        }

        // 現在のキーワードを保存（異常終了対策）
        saveCurrentKeyword()

        // 検索履歴に追加
        addToHistory(pattern)

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
        statusMessage = "検索準備中..."
        log.debug("バックグラウンドタスク開始")

        let onlySentCopy = onlySent
        currentSearchTask = Task.detached { [weak self] in
            await self?.performSearch(searchPattern: searchPattern, onlySent: onlySentCopy)
        }
    }

    /// 検索をキャンセル
    func cancelSearch() {
        if let task = currentSearchTask {
            log.info("検索キャンセル")
            task.cancel()
            currentSearchTask = nil
            isSearching = false
            searchCompleted = false
            statusMessage = "検索がキャンセルされました"
        }
    }

    private nonisolated func performSearch(searchPattern: SearchPattern, onlySent: Bool) async {
        let searchLog = CategoryLogger(category: .search)
        searchLog.info("performSearch開始: 送信済みのみ=\(onlySent)")

        let localMailFolderService = MailFolderService()
        let localEmlxParser = EmlxParser()

        searchLog.debug("emlxファイル収集中...")

        await MainActor.run { [weak self] in
            self?.statusMessage = "メールファイルを収集中..."
        }

        // キャンセルチェック
        if Task.isCancelled {
            searchLog.info("検索がキャンセルされました（ファイル収集前）")
            return
        }

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
        var lastUpdateCount = 0

        for (index, emlxURL) in emlxFiles.enumerated() {
            // キャンセルチェック（各ファイル処理前）
            if Task.isCancelled {
                searchLog.info("検索がキャンセルされました（\(index)/\(total)）")
                await MainActor.run { [weak self] in
                    self?.isSearching = false
                    self?.searchCompleted = false
                    self?.statusMessage = "検索がキャンセルされました"
                    self?.currentSearchTask = nil
                }
                return
            }

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

            // 進捗更新とリアルタイム結果更新（100件ごと、または新しいヒットがあるとき）
            let shouldUpdate = (index % 100 == 0) || (index == total - 1) || (hitLines.count > lastUpdateCount)
            if shouldUpdate {
                let currentProgress = Double(index + 1) / Double(total)
                let currentHits = hitLines.count
                let currentResults = hitLines  // コピーを作成
                lastUpdateCount = currentHits

                await MainActor.run { [weak self] in
                    self?.progress = currentProgress
                    self?.results = currentResults  // リアルタイムで結果を更新
                    self?.statusMessage = "\(index + 1)/\(total) 検索中... (\(currentHits)件ヒット)"
                }
            }
        }

        // 検索完了
        let finalResults = hitLines
        let finalMailCount = Set(hitLines.map(\.profile.messageID)).count

        searchLog.info("検索完了: \(finalResults.count)件ヒット（\(finalMailCount)通）、パースエラー: \(parseErrors)件")

        await MainActor.run { [weak self] in
            self?.results = finalResults
            self?.isSearching = false
            self?.searchCompleted = true
            self?.statusMessage = "検索完了: \(finalResults.count)件ヒット（\(finalMailCount)通）"
            self?.progress = 1.0
            self?.currentSearchTask = nil
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
        var csv = "\u{FEFF}" // BOM for UTF-8
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
