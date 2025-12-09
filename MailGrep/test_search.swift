#!/usr/bin/env swift

import Foundation

// MailFolderServiceの簡易版
func collectEmlxFiles() -> [URL] {
    let fileManager = FileManager.default
    let homeDir = fileManager.homeDirectoryForCurrentUser
    let mailDir = homeDir.appendingPathComponent("Library/Mail")

    print("メールディレクトリ: \(mailDir.path)")

    guard fileManager.fileExists(atPath: mailDir.path) else {
        print("エラー: メールディレクトリが存在しません")
        return []
    }

    var emlxFiles: [URL] = []

    if let enumerator = fileManager.enumerator(
        at: mailDir,
        includingPropertiesForKeys: [.isRegularFileKey],
        options: [.skipsHiddenFiles]
    ) {
        for case let fileURL as URL in enumerator {
            if fileURL.pathExtension == "emlx" {
                emlxFiles.append(fileURL)
            }
        }
    }

    return emlxFiles
}

// テスト実行
print("=== 検索機能テスト開始 ===\n")

// 1. emlxファイルを収集
print("1. emlxファイルを収集中...")
let emlxFiles = collectEmlxFiles()
print("   結果: \(emlxFiles.count)件のemlxファイルを発見\n")

if emlxFiles.count > 0 {
    print("   最初の5件:")
    for (index, file) in emlxFiles.prefix(5).enumerated() {
        print("   [\(index + 1)] \(file.lastPathComponent)")
    }
    print("")
}

// 2. 最初のemlxファイルを読み込んでみる
if let firstFile = emlxFiles.first {
    print("2. 最初のemlxファイルを読み込み中...")
    print("   ファイル: \(firstFile.path)")

    do {
        let content = try String(contentsOf: firstFile, encoding: .utf8)
        let lines = content.components(separatedBy: .newlines)
        print("   行数: \(lines.count)")
        print("   最初の10行:")
        for (index, line) in lines.prefix(10).enumerated() {
            let truncated = line.count > 80 ? String(line.prefix(80)) + "..." : line
            print("   [\(index + 1)] \(truncated)")
        }
    } catch {
        print("   エラー: \(error.localizedDescription)")
    }
    print("")
}

// 3. 正規表現テスト
print("3. 正規表現テスト...")
let pattern = "test"
do {
    let regex = try NSRegularExpression(pattern: pattern, options: [.caseInsensitive])
    let testString = "This is a Test string"
    let range = NSRange(testString.startIndex..., in: testString)
    let matches = regex.matches(in: testString, options: [], range: range)
    print("   パターン: '\(pattern)'")
    print("   テスト文字列: '\(testString)'")
    print("   マッチ数: \(matches.count)")
} catch {
    print("   正規表現エラー: \(error.localizedDescription)")
}

print("\n=== テスト完了 ===")
