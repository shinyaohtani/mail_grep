import AppKit
import Foundation

class MailLinkService {

    func openInMailApp(messageID: String) {
        guard !messageID.isEmpty else { return }

        let cleanID = messageID.trimmingCharacters(in: CharacterSet(charactersIn: "<>"))

        guard let encoded = cleanID.addingPercentEncoding(withAllowedCharacters: .urlPathAllowed),
              let url = URL(string: "message:\(encoded)") else { return }

        NSWorkspace.shared.open(url)
    }

    func openInMailApp(profile: MailProfile) {
        openInMailApp(messageID: profile.messageID)
    }

    func openInMailApp(hitLine: HitLine) {
        openInMailApp(profile: hitLine.profile)
    }
}
