import AppKit
import Foundation

class MailLinkService {
    /// Message-IDからMail.appで開くためのURLを生成
    /// Python版と同じ形式: message:%3Cxxx%40yyy.zzz%3E
    func generateMailLink(messageID: String) -> String? {
        guard !messageID.isEmpty else { return nil }

        var id = messageID.trimmingCharacters(in: .whitespaces)

        // 角括弧で囲まれていなければ追加（Python版と同じ処理）
        if !id.hasPrefix("<") {
            id = "<\(id)>"
        }

        // Python版のquote(s, safe='')と同様に、全ての特殊文字をエンコード
        // 空のCharacterSetを使うと全ての文字がエンコードされる
        guard let encoded = id.addingPercentEncoding(withAllowedCharacters: CharacterSet()) else {
            return nil
        }

        return "message:\(encoded)"
    }

    func openInMailApp(messageID: String) {
        guard let link = generateMailLink(messageID: messageID),
              let url = URL(string: link) else { return }

        NSWorkspace.shared.open(url)
    }

    func openInMailApp(profile: MailProfile) {
        openInMailApp(messageID: profile.messageID)
    }

    func openInMailApp(hitLine: HitLine) {
        openInMailApp(profile: hitLine.profile)
    }

    /// Excel用のHYPERLINK関数形式でリンクを生成
    /// Python版と同じ形式: =HYPERLINK("message:%3Cxxx%40yyy.zzz%3E","メール")
    func generateExcelLink(messageID: String) -> String {
        guard let link = generateMailLink(messageID: messageID) else {
            return ""
        }
        return "=HYPERLINK(\"\(link)\",\"メール\")"
    }
}
