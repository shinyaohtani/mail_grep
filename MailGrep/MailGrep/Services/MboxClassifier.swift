import Foundation

private let log = CategoryLogger(category: .classifier)

class MboxClassifier {
    private let excludeTokens: Set<String> = [
        // drafts
        "draft", "drafts", "下書", "brouillon", "bozza", "entwurf", "borrador", "rascunho", "черновик",
        // trash / deleted / bin
        "trash", "deleted", "bin", "ゴミ箱", "削除済", "corbeille", "papierkorb", "papelera", "cestino", "корзина",
        // junk / spam
        "junk", "spam", "迷惑メール", "indesirable", "indésirable", "unerwünscht", "спам",
        // outbox
        "outbox", "送信トレイ", "posteausgang",
        // archive
        "archive", "アーカイブ", "archivo", "archivio", "archiv",
        // noise
        "rss", "メモ", "notes", "tasks", "タスク", "journal", "会話の履歴", "同期の問題", "recovered",
    ]

    private let sentTokens: Set<String> = [
        "sent", "送信済み", "送信済みアイテム", "gesendet", "inviati", "enviados", "envoyes", "envoyés", "отправленные",
    ]

    private let specialAttrTokens: Set<String> = [
        "\\drafts", "\\junk", "\\trash", "\\deleted", "\\bin", "\\spam", "\\outbox", "\\archive",
    ]

    func isExcluded(_ mboxDir: URL) -> Bool {
        let infoPath = mboxDir.appendingPathComponent("Info.plist")
        let mboxName = mboxDir.lastPathComponent

        if FileManager.default.fileExists(atPath: infoPath.path) {
            let strings = stringsFromPlist(infoPath)
            log.debug("isExcluded判定: \(mboxName) → plistから\(strings.count)個の文字列")

            if let matchedToken = hitToken(specialAttrTokens, in: strings) {
                log.debug("\(mboxName) → 特殊属性で除外: \(matchedToken)")
                return true
            }
            if let matchedToken = hitToken(excludeTokens, in: strings) {
                log.debug("\(mboxName) → plistトークンで除外: \(matchedToken)")
                return true
            }
        }

        let names = [
            mboxDir.lastPathComponent.replacingOccurrences(of: ".mbox", with: ""),
            mboxDir.deletingLastPathComponent().lastPathComponent,
        ]
        if let matchedToken = hitToken(excludeTokens, in: names) {
            log.debug("\(mboxName) → フォルダ名トークンで除外: \(matchedToken)")
            return true
        }
        return false
    }

    func isSent(_ mboxDir: URL) -> Bool {
        let infoPath = mboxDir.appendingPathComponent("Info.plist")

        if FileManager.default.fileExists(atPath: infoPath.path) {
            let strings = stringsFromPlist(infoPath)

            if hit(["\\sent"], in: strings) {
                return true
            }
            if hit(sentTokens, in: strings) {
                return true
            }
        }

        let names = [
            mboxDir.lastPathComponent.replacingOccurrences(of: ".mbox", with: ""),
            mboxDir.deletingLastPathComponent().lastPathComponent,
        ]
        return hit(sentTokens, in: names)
    }

    private func normalize(_ s: String) -> String {
        let nfkc = s.precomposedStringWithCompatibilityMapping.lowercased()
        let pattern = "[\\s\\W_]+"
        guard let regex = try? NSRegularExpression(pattern: pattern, options: [.useUnicodeWordBoundaries]) else {
            return nfkc
        }
        return regex.stringByReplacingMatches(in: nfkc, range: NSRange(nfkc.startIndex..., in: nfkc), withTemplate: "")
    }

    /// 分類に必要なキーのみをホワイトリストで抽出
    /// （全文字列を抽出すると、ExchangeSyncState等のバイナリデータから偽陽性が発生するため）
    private let classificationKeys: Set<String> = [
        "MailboxName", // 表示名（受信トレイ、送信済みアイテム等）
        "IMAPMailboxName", // IMAPフォルダ名
        "SpecialMailboxType", // 特殊メールボックスタイプ
        "AccountPath", // アカウントパス内のフォルダ名
        "CriteriaCriteria", // スマートフォルダの条件
    ]

    private func stringsFromPlist(_ path: URL) -> [String] {
        var result: [String] = []

        guard let data = try? Data(contentsOf: path),
              let plist = try? PropertyListSerialization.propertyList(from: data, format: nil) as? [String: Any]
        else {
            return result
        }

        // ホワイトリストに含まれるキーの値のみを抽出
        for key in classificationKeys {
            if let value = plist[key] as? String {
                result.append(value)
            }
        }

        return result
    }

    private func hit(_ tokens: Set<String>, in candidates: [String]) -> Bool {
        hitToken(tokens, in: candidates) != nil
    }

    private func hitToken(_ tokens: Set<String>, in candidates: [String]) -> String? {
        for candidate in candidates {
            let normalized = normalize(candidate)
            for token in tokens {
                let normalizedToken = normalize(token)
                if normalized.contains(normalizedToken) {
                    return token
                }
            }
        }
        return nil
    }
}
