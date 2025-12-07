import Foundation

struct HitLine: Identifiable, Hashable {
    let id = UUID()
    let mailID: Int
    let profile: MailProfile
    let matchedLine: String
    let lineNumber: Int

    var dateStr: String { profile.dateStr }
    var subject: String { profile.subject }
    var fromAddr: String { profile.fromAddr }
    var toAddr: String { profile.toAddr }

    func csvRow() -> String {
        let sanitizedLine = matchedLine
            .replacingOccurrences(of: "\"", with: "\"\"")
            .replacingOccurrences(of: "\n", with: " ")
            .replacingOccurrences(of: "\r", with: " ")
        let sanitizedSubject = subject
            .replacingOccurrences(of: "\"", with: "\"\"")
        return "\(mailID),\"\(dateStr)\",\"\(sanitizedSubject)\",\"\(fromAddr)\",\"\(toAddr)\",\"\(sanitizedLine)\""
    }
}
