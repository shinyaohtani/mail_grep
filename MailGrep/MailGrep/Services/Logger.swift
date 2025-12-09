// Logger.swift
// Python版 SmartLogging に対応するSwift版ロギングシステム
//
// 機能:
// - ログレベル制御（debug, info, warning, error）
// - カテゴリ別ロガー（ViewModel, Parser, Service等）
// - 経過時間計測
// - ファイル出力（オプション）
// - 長文の切り詰め（truncate）

import Foundation
import os

// MARK: - Log Level

enum LogLevel: Int, Comparable {
    case debug = 0
    case info = 1
    case warning = 2
    case error = 3

    static func < (lhs: LogLevel, rhs: LogLevel) -> Bool {
        lhs.rawValue < rhs.rawValue
    }

    var osLogType: OSLogType {
        switch self {
        case .debug: .debug
        case .info: .info
        case .warning: .default
        case .error: .error
        }
    }

    var prefix: String {
        switch self {
        case .debug: "🔍 DEBUG"
        case .info: "ℹ️ INFO"
        case .warning: "⚠️ WARN"
        case .error: "❌ ERROR"
        }
    }
}

// MARK: - Logger Category

enum LogCategory: String {
    case app = "App"
    case viewModel = "ViewModel"
    case parser = "Parser"
    case service = "Service"
    case classifier = "Classifier"
    case search = "Search"

    var subsystem: String { "com.mailgrep.app" }
}

// MARK: - Smart Logger

final class SmartLogger {
    static let shared = SmartLogger()

    private let startTime: Date
    private var currentLevel: LogLevel = .info
    private var fileHandle: FileHandle?
    private var logFileURL: URL?
    private let queue = DispatchQueue(label: "com.mailgrep.logger", qos: .utility)

    private init() {
        startTime = Date()
        setupFileLogging()
    }

    // MARK: - Configuration

    /// 環境変数 MAIL_GREP_DEBUG=1 でデバッグモード有効
    var isDebugMode: Bool {
        ProcessInfo.processInfo.environment["MAIL_GREP_DEBUG"] == "1"
    }

    /// ログレベルを設定
    func setLevel(_ level: LogLevel) {
        currentLevel = level
    }

    /// 実効ログレベル（デバッグモードなら常にdebug）
    var effectiveLevel: LogLevel {
        isDebugMode ? .debug : currentLevel
    }

    // MARK: - File Logging Setup

    private func setupFileLogging() {
        guard let logsDir = FileManager.default.urls(for: .libraryDirectory, in: .userDomainMask).first?
            .appendingPathComponent("Logs")
            .appendingPathComponent("MailGrep") else { return }

        do {
            try FileManager.default.createDirectory(at: logsDir, withIntermediateDirectories: true)
            let dateFormatter = DateFormatter()
            dateFormatter.dateFormat = "yyyyMMdd_HHmmss"
            let filename = "mailgrep_\(dateFormatter.string(from: startTime)).log"
            logFileURL = logsDir.appendingPathComponent(filename)

            if let url = logFileURL {
                FileManager.default.createFile(atPath: url.path, contents: nil)
                fileHandle = try FileHandle(forWritingTo: url)
            }
        } catch {
            // ファイルログ設定失敗は無視（コンソールログは継続）
        }
    }

    // MARK: - Logging Methods

    func debug(_ message: @autoclosure () -> String, category: LogCategory = .app, file: String = #file, function: String = #function, line: Int = #line) {
        log(level: .debug, message: message(), category: category, file: file, function: function, line: line)
    }

    func info(_ message: @autoclosure () -> String, category: LogCategory = .app, file: String = #file, function: String = #function, line: Int = #line) {
        log(level: .info, message: message(), category: category, file: file, function: function, line: line)
    }

    func warning(_ message: @autoclosure () -> String, category: LogCategory = .app, file: String = #file, function: String = #function, line: Int = #line) {
        log(level: .warning, message: message(), category: category, file: file, function: function, line: line)
    }

    func error(_ message: @autoclosure () -> String, category: LogCategory = .app, file: String = #file, function: String = #function, line: Int = #line) {
        log(level: .error, message: message(), category: category, file: file, function: function, line: line)
    }

    private func log(level: LogLevel, message: String, category: LogCategory, file: String, function _: String, line: Int) {
        guard level >= effectiveLevel else { return }

        let elapsed = elapsedTimeString()
        let fileName = URL(fileURLWithPath: file).deletingPathExtension().lastPathComponent
        let locationInfo = isDebugMode ? " [\(fileName):\(line)]" : ""
        let formattedMessage = "[\(elapsed)] \(level.prefix) [\(category.rawValue)]\(locationInfo) \(message)"

        // Console output
        let osLogger = os.Logger(subsystem: category.subsystem, category: category.rawValue)
        osLogger.log(level: level.osLogType, "\(formattedMessage)")

        // File output (async)
        queue.async { [weak self] in
            self?.writeToFile(formattedMessage)
        }
    }

    private func writeToFile(_ message: String) {
        guard let handle = fileHandle,
              let data = (message + "\n").data(using: .utf8) else { return }
        do {
            try handle.write(contentsOf: data)
        } catch {
            // ファイル書き込みエラーは無視（クローズ後の書き込み試行など）
        }
    }

    // MARK: - Utilities

    /// 経過時間を h:mm:ss 形式で返す
    func elapsedTimeString() -> String {
        let elapsed = Date().timeIntervalSince(startTime)
        let hours = Int(elapsed) / 3600
        let minutes = (Int(elapsed) % 3600) / 60
        let seconds = Int(elapsed) % 60
        return String(format: "%d:%02d:%02d", hours, minutes, seconds)
    }

    /// 経過時間（秒）
    var elapsedSeconds: TimeInterval {
        Date().timeIntervalSince(startTime)
    }

    /// 文字列を指定長に切り詰め
    static func truncate(_ text: String, maxLength: Int = 50) -> String {
        guard text.count > maxLength else { return text }
        let index = text.index(text.startIndex, offsetBy: maxLength - 3)
        return String(text[..<index]) + "..."
    }

    /// ログファイルのパスを返す
    var logFilePath: String? {
        logFileURL?.path
    }

    /// 終了処理（ファイルクローズ＆サマリー出力）
    func finalize() {
        let elapsed = elapsedTimeString()
        info("検索完了 - 経過時間: \(elapsed)", category: .app)
        if let path = logFilePath {
            info("ログファイル: \(path)", category: .app)
        }
        queue.sync {
            try? fileHandle?.close()
            fileHandle = nil
        }
    }

    deinit {
        try? fileHandle?.close()
    }
}

// MARK: - Convenience Global Access

/// グローバルロガーインスタンス
let Log = SmartLogger.shared

// MARK: - Category-Specific Loggers

/// カテゴリ固定のロガーラッパー
struct CategoryLogger {
    let category: LogCategory

    func debug(_ message: @autoclosure () -> String, file: String = #file, function: String = #function, line: Int = #line) {
        Log.debug(message(), category: category, file: file, function: function, line: line)
    }

    func info(_ message: @autoclosure () -> String, file: String = #file, function: String = #function, line: Int = #line) {
        Log.info(message(), category: category, file: file, function: function, line: line)
    }

    func warning(_ message: @autoclosure () -> String, file: String = #file, function: String = #function, line: Int = #line) {
        Log.warning(message(), category: category, file: file, function: function, line: line)
    }

    func error(_ message: @autoclosure () -> String, file: String = #file, function: String = #function, line: Int = #line) {
        Log.error(message(), category: category, file: file, function: function, line: line)
    }
}

// MARK: - Scoped Logging (for timing)

/// スコープ付きロギング（処理時間計測用）
final class ScopedLog {
    private let message: String
    private let category: LogCategory
    private let startTime: Date
    private let file: String
    private let function: String
    private let line: Int

    init(_ message: String, category: LogCategory = .app, file: String = #file, function: String = #function, line: Int = #line) {
        self.message = message
        self.category = category
        startTime = Date()
        self.file = file
        self.function = function
        self.line = line
        Log.debug("\(message) 開始", category: category, file: file, function: function, line: line)
    }

    deinit {
        let elapsed = Date().timeIntervalSince(startTime)
        let elapsedStr = String(format: "%.2f秒", elapsed)
        Log.debug("\(message) 完了 (\(elapsedStr))", category: category, file: file, function: function, line: line)
    }
}
