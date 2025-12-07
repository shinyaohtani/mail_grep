import Foundation

struct ParsedMail {
    let headers: MailHeaders
    let bodyLines: [(line: String, type: String)]
}

struct MailHeaders {
    let messageID: String
    let dateStr: String
    let dateValue: Date?
    let subject: String
    let from: String
    let to: String

    var headerLines: [String] {
        var lines: [String] = []
        if !subject.isEmpty { lines.append("Subject: \(subject)") }
        if !from.isEmpty { lines.append("From: \(from)") }
        if !to.isEmpty { lines.append("To: \(to)") }
        if !dateStr.isEmpty { lines.append("Date: \(dateStr)") }
        return lines
    }
}

class EmlxParser {
    private let headerDecoder = HeaderDecoder()

    func parse(url: URL) throws -> ParsedMail {
        let rawData = try Data(contentsOf: url)
        let data = stripByteSizePrefix(rawData)

        guard let headerEndIndex = findHeaderEnd(in: data) else {
            throw EmlxParseError.invalidFormat
        }

        let headerData = data[..<headerEndIndex]
        let bodyData = data[headerEndIndex...]

        let headers = parseHeaders(headerData)
        let bodyLines = parseBody(bodyData, headers: headers)

        return ParsedMail(headers: headers, bodyLines: bodyLines)
    }

    private func stripByteSizePrefix(_ data: Data) -> Data {
        guard let firstByte = data.first, firstByte >= 0x30, firstByte <= 0x39 else {
            return data
        }
        if let newlineIndex = data.firstIndex(of: 0x0A) {
            return data[(newlineIndex + 1)...]
        }
        return data
    }

    private func findHeaderEnd(in data: Data) -> Data.Index? {
        let doubleCRLF = Data([0x0D, 0x0A, 0x0D, 0x0A])
        let doubleLF = Data([0x0A, 0x0A])

        if let range = data.range(of: doubleCRLF) {
            return range.upperBound
        }
        if let range = data.range(of: doubleLF) {
            return range.upperBound
        }
        return nil
    }

    private func parseHeaders(_ data: Data) -> MailHeaders {
        let text = decodeHeaderData(data)
        let lines = unfoldHeaders(text)
        var headersDict: [String: String] = [:]

        for line in lines {
            if let colonIndex = line.firstIndex(of: ":") {
                let key = String(line[..<colonIndex]).trimmingCharacters(in: .whitespaces).lowercased()
                let value = String(line[line.index(after: colonIndex)...]).trimmingCharacters(in: .whitespaces)
                headersDict[key] = value
            }
        }

        let messageID = headersDict["message-id"] ?? ""
        let rawDate = headersDict["date"] ?? ""
        let (dateStr, dateValue) = parseDate(rawDate)

        return MailHeaders(
            messageID: headerDecoder.decode(messageID),
            dateStr: dateStr,
            dateValue: dateValue,
            subject: headerDecoder.decode(headersDict["subject"] ?? ""),
            from: headerDecoder.decode(headersDict["from"] ?? ""),
            to: headerDecoder.decode(headersDict["to"] ?? "")
        )
    }

    private func decodeHeaderData(_ data: Data) -> String {
        if let text = String(data: Data(data), encoding: .utf8) {
            return text
        }
        if let text = String(data: Data(data), encoding: .isoLatin1) {
            return text
        }
        return String(data: Data(data), encoding: .ascii) ?? ""
    }

    private func unfoldHeaders(_ text: String) -> [String] {
        var result: [String] = []
        var current = ""

        for line in text.components(separatedBy: .newlines) {
            if line.isEmpty { break }
            if line.first?.isWhitespace == true {
                current += " " + line.trimmingCharacters(in: .whitespaces)
            } else {
                if !current.isEmpty {
                    result.append(current)
                }
                current = line
            }
        }
        if !current.isEmpty {
            result.append(current)
        }
        return result
    }

    private func parseDate(_ rawDate: String) -> (String, Date?) {
        let dateFormats = [
            "EEE, dd MMM yyyy HH:mm:ss Z",
            "EEE, dd MMM yyyy HH:mm:ss ZZZZZ",
            "dd MMM yyyy HH:mm:ss Z",
            "EEE, d MMM yyyy HH:mm:ss Z",
            "d MMM yyyy HH:mm:ss Z"
        ]

        let formatter = DateFormatter()
        formatter.locale = Locale(identifier: "en_US_POSIX")

        let cleanDate = rawDate.replacingOccurrences(of: "\\s*\\([^)]*\\)", with: "", options: .regularExpression)

        for format in dateFormats {
            formatter.dateFormat = format
            if let date = formatter.date(from: cleanDate.trimmingCharacters(in: .whitespaces)) {
                let outputFormatter = DateFormatter()
                outputFormatter.dateFormat = "yyyy-MM-dd HH:mm:ss"
                return (outputFormatter.string(from: date), date)
            }
        }

        return ("", nil)
    }

    private func parseBody(_ data: Data, headers: MailHeaders) -> [(String, String)] {
        var result: [(String, String)] = []

        let contentType = extractContentType(from: data) ?? "text/plain"
        let charset = extractCharset(from: data) ?? "utf-8"

        if contentType.hasPrefix("multipart/") {
            if let boundary = extractBoundary(from: data) {
                result = parseMultipart(data, boundary: boundary)
            }
        } else {
            let text = decodeBody(data, charset: charset)
            for line in text.components(separatedBy: .newlines) where !line.trimmingCharacters(in: .whitespaces).isEmpty {
                result.append((line, "text/plain"))
            }
        }

        return result
    }

    private func extractContentType(from data: Data) -> String? {
        guard let text = String(data: Data(data.prefix(2000)), encoding: .ascii) else { return nil }
        let pattern = "Content-Type:\\s*([^;\\r\\n]+)"
        guard let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
              let match = regex.firstMatch(in: text, range: NSRange(text.startIndex..., in: text)),
              let range = Range(match.range(at: 1), in: text) else { return nil }
        return String(text[range]).trimmingCharacters(in: .whitespaces).lowercased()
    }

    private func extractCharset(from data: Data) -> String? {
        guard let text = String(data: Data(data.prefix(2000)), encoding: .ascii) else { return nil }
        let pattern = "charset=[\"]?([^\"\\s;]+)"
        guard let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
              let match = regex.firstMatch(in: text, range: NSRange(text.startIndex..., in: text)),
              let range = Range(match.range(at: 1), in: text) else { return nil }
        return String(text[range]).lowercased()
    }

    private func extractBoundary(from data: Data) -> String? {
        guard let text = String(data: Data(data.prefix(2000)), encoding: .ascii) else { return nil }
        let pattern = "boundary=[\"]?([^\"\\s;\\r\\n]+)"
        guard let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
              let match = regex.firstMatch(in: text, range: NSRange(text.startIndex..., in: text)),
              let range = Range(match.range(at: 1), in: text) else { return nil }
        return String(text[range])
    }

    private func parseMultipart(_ data: Data, boundary: String) -> [(String, String)] {
        var result: [(String, String)] = []
        guard let text = String(data: Data(data), encoding: .utf8) ?? String(data: Data(data), encoding: .isoLatin1) else {
            return result
        }

        let delimiter = "--\(boundary)"
        let parts = text.components(separatedBy: delimiter)

        for part in parts {
            if part.hasPrefix("--") || part.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty {
                continue
            }

            let partContentType = extractPartContentType(part) ?? "text/plain"
            let partCharset = extractPartCharset(part) ?? "utf-8"

            if partContentType.hasPrefix("text/") {
                let bodyStart = part.range(of: "\r\n\r\n") ?? part.range(of: "\n\n")
                if let start = bodyStart {
                    let bodyText = String(part[start.upperBound...])
                    let decoded = decodePartBody(bodyText, charset: partCharset, part: part)

                    if partContentType == "text/html" {
                        let plainText = stripHTML(decoded)
                        for line in plainText.components(separatedBy: .newlines)
                        where !line.trimmingCharacters(in: .whitespaces).isEmpty {
                            result.append((line, "text/html_textonly"))
                        }
                    } else {
                        for line in decoded.components(separatedBy: .newlines)
                        where !line.trimmingCharacters(in: .whitespaces).isEmpty {
                            result.append((line, partContentType))
                        }
                    }
                }
            }
        }

        return result
    }

    private func extractPartContentType(_ part: String) -> String? {
        let pattern = "Content-Type:\\s*([^;\\r\\n]+)"
        guard let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
              let match = regex.firstMatch(in: part, range: NSRange(part.startIndex..., in: part)),
              let range = Range(match.range(at: 1), in: part) else { return nil }
        return String(part[range]).trimmingCharacters(in: .whitespaces).lowercased()
    }

    private func extractPartCharset(_ part: String) -> String? {
        let pattern = "charset=[\"]?([^\"\\s;\\r\\n]+)"
        guard let regex = try? NSRegularExpression(pattern: pattern, options: .caseInsensitive),
              let match = regex.firstMatch(in: part, range: NSRange(part.startIndex..., in: part)),
              let range = Range(match.range(at: 1), in: part) else { return nil }
        return String(part[range]).lowercased()
    }

    private func decodePartBody(_ text: String, charset: String, part: String) -> String {
        let encodingPattern = "Content-Transfer-Encoding:\\s*([^\\s\\r\\n]+)"
        var encoding = "7bit"
        if let regex = try? NSRegularExpression(pattern: encodingPattern, options: .caseInsensitive),
           let match = regex.firstMatch(in: part, range: NSRange(part.startIndex..., in: part)),
           let range = Range(match.range(at: 1), in: part) {
            encoding = String(part[range]).lowercased()
        }

        switch encoding {
        case "base64":
            let cleaned = text.replacingOccurrences(of: "\\s", with: "", options: .regularExpression)
            if let data = Data(base64Encoded: cleaned) {
                return decodeBody(data, charset: charset)
            }
        case "quoted-printable":
            return decodeQuotedPrintable(text, charset: charset)
        default:
            break
        }

        return text
    }

    private func decodeBody(_ data: Data, charset: String) -> String {
        let encoding = charsetToEncoding(charset)
        if let text = String(data: data, encoding: encoding) {
            return text
        }
        return String(data: data, encoding: .utf8) ?? String(data: data, encoding: .isoLatin1) ?? ""
    }

    private func decodeQuotedPrintable(_ text: String, charset: String) -> String {
        var result = text
        result = result.replacingOccurrences(of: "=\r\n", with: "")
        result = result.replacingOccurrences(of: "=\n", with: "")

        let pattern = "=([0-9A-Fa-f]{2})"
        guard let regex = try? NSRegularExpression(pattern: pattern) else { return result }

        var output = ""
        var bytes: [UInt8] = []
        var lastEnd = result.startIndex

        let matches = regex.matches(in: result, range: NSRange(result.startIndex..., in: result))
        for match in matches {
            guard let range = Range(match.range, in: result),
                  let hexRange = Range(match.range(at: 1), in: result) else { continue }

            output += result[lastEnd..<range.lowerBound]
            if !bytes.isEmpty {
                let data = Data(bytes)
                output += decodeBody(data, charset: charset)
                bytes = []
            }

            if let byte = UInt8(String(result[hexRange]), radix: 16) {
                bytes.append(byte)
            }
            lastEnd = range.upperBound
        }

        if !bytes.isEmpty {
            let data = Data(bytes)
            output += decodeBody(data, charset: charset)
        }
        output += result[lastEnd...]

        return output
    }

    private func charsetToEncoding(_ charset: String) -> String.Encoding {
        switch charset.lowercased() {
        case "utf-8", "utf8":
            return .utf8
        case "iso-2022-jp":
            return .iso2022JP
        case "shift_jis", "shift-jis", "sjis", "x-sjis":
            return .shiftJIS
        case "euc-jp", "eucjp":
            return .japaneseEUC
        case "iso-8859-1", "latin1":
            return .isoLatin1
        case "us-ascii", "ascii":
            return .ascii
        default:
            return .utf8
        }
    }

    private func stripHTML(_ html: String) -> String {
        var result = html

        let tagsToRemove = ["<head[^>]*>.*?</head>", "<script[^>]*>.*?</script>", "<style[^>]*>.*?</style>"]
        for pattern in tagsToRemove {
            if let regex = try? NSRegularExpression(pattern: pattern, options: [.caseInsensitive, .dotMatchesLineSeparators]) {
                result = regex.stringByReplacingMatches(in: result, range: NSRange(result.startIndex..., in: result), withTemplate: "")
            }
        }

        if let regex = try? NSRegularExpression(pattern: "<[^>]+>", options: []) {
            result = regex.stringByReplacingMatches(in: result, range: NSRange(result.startIndex..., in: result), withTemplate: " ")
        }

        let entities = [("&nbsp;", " "), ("&lt;", "<"), ("&gt;", ">"), ("&amp;", "&"), ("&quot;", "\"")]
        for (entity, char) in entities {
            result = result.replacingOccurrences(of: entity, with: char)
        }

        result = result.replacingOccurrences(of: "\\s+", with: " ", options: .regularExpression)

        return result.trimmingCharacters(in: .whitespacesAndNewlines)
    }
}

enum EmlxParseError: Error {
    case invalidFormat
    case encodingError
}
