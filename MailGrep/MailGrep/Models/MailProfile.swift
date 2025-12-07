import Foundation

struct MailProfile: Identifiable, Hashable {
    let id = UUID()
    let messageID: String
    let dateStr: String
    let dateValue: Date?
    let link: String
    let subject: String
    let fromAddr: String
    let toAddr: String
    let emlxPath: URL

    var mailAppURL: URL? {
        guard !messageID.isEmpty else { return nil }
        let cleanID = messageID
            .trimmingCharacters(in: CharacterSet(charactersIn: "<>"))
        guard let encoded = cleanID.addingPercentEncoding(withAllowedCharacters: .urlPathAllowed) else {
            return nil
        }
        return URL(string: "message:\(encoded)")
    }
}
