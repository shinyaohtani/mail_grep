import Foundation

struct SearchPattern {
    let pattern: String
    let regex: NSRegularExpression
    let ignoreCase: Bool

    init(pattern: String, ignoreCase: Bool = true) throws {
        self.pattern = pattern
        self.ignoreCase = ignoreCase

        var options: NSRegularExpression.Options = []
        if ignoreCase {
            options.insert(.caseInsensitive)
        }

        self.regex = try NSRegularExpression(pattern: pattern, options: options)
    }

    func matches(_ line: String) -> Bool {
        let range = NSRange(line.startIndex..., in: line)
        return regex.firstMatch(in: line, options: [], range: range) != nil
    }

    var uniqueName: String {
        let cleaned = pattern
            .replacingOccurrences(of: "[^\\p{L}\\p{N}\\s]", with: "", options: .regularExpression)
            .replacingOccurrences(of: "\\s+", with: "_", options: .regularExpression)

        let prefix = cleaned.isEmpty ? "search" : String(cleaned.prefix(16))
        let timestamp = ISO8601DateFormatter().string(from: Date())
            .replacingOccurrences(of: ":", with: "")
            .replacingOccurrences(of: "-", with: "")

        return "\(prefix)_\(timestamp)"
    }
}
