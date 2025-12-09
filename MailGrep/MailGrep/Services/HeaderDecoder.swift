import Foundation

class HeaderDecoder {
    func decode(_ value: String) -> String {
        guard !value.isEmpty else { return "" }

        let cleaned = removeCRLF(value)
        let pattern = "=\\?([^?]+)\\?([BQbq])\\?([^?]*)\\?="

        guard let regex = try? NSRegularExpression(pattern: pattern) else {
            return cleaned
        }

        var result = ""
        var lastEnd = cleaned.startIndex
        let matches = regex.matches(in: cleaned, range: NSRange(cleaned.startIndex..., in: cleaned))

        for match in matches {
            guard let fullRange = Range(match.range, in: cleaned),
                  let charsetRange = Range(match.range(at: 1), in: cleaned),
                  let encodingRange = Range(match.range(at: 2), in: cleaned),
                  let textRange = Range(match.range(at: 3), in: cleaned) else { continue }

            result += cleaned[lastEnd ..< fullRange.lowerBound]

            let charset = String(cleaned[charsetRange])
            let encoding = String(cleaned[encodingRange]).uppercased()
            let encodedText = String(cleaned[textRange])

            let decoded = decodeEncodedWord(text: encodedText, encoding: encoding, charset: charset)
            result += decoded

            lastEnd = fullRange.upperBound
        }

        result += cleaned[lastEnd...]
        return result
    }

    private func decodeEncodedWord(text: String, encoding: String, charset: String) -> String {
        var data: Data?

        switch encoding {
        case "B":
            data = Data(base64Encoded: text)
        case "Q":
            data = decodeQuotedPrintableHeader(text)
        default:
            return text
        }

        guard let decodedData = data else { return text }

        let stringEncoding = charsetToEncoding(charset)
        if let decoded = String(data: decodedData, encoding: stringEncoding) {
            return decoded
        }
        if let decoded = String(data: decodedData, encoding: .utf8) {
            return decoded
        }
        if let decoded = String(data: decodedData, encoding: .isoLatin1) {
            return decoded
        }

        return text
    }

    private func decodeQuotedPrintableHeader(_ text: String) -> Data? {
        var bytes: [UInt8] = []
        var index = text.startIndex

        while index < text.endIndex {
            let char = text[index]

            if char == "=" {
                let nextIndex = text.index(after: index)
                if nextIndex < text.endIndex,
                   let hexEndIndex = text.index(nextIndex, offsetBy: 2, limitedBy: text.endIndex)
                {
                    let hex = String(text[nextIndex ..< hexEndIndex])
                    if hex.count == 2, let byte = UInt8(hex, radix: 16) {
                        bytes.append(byte)
                        index = hexEndIndex
                        continue
                    }
                }
                bytes.append(UInt8(ascii: "="))
            } else if char == "_" {
                bytes.append(0x20)
            } else {
                for byte in char.utf8 {
                    bytes.append(byte)
                }
            }
            index = text.index(after: index)
        }

        return Data(bytes)
    }

    private func charsetToEncoding(_ charset: String) -> String.Encoding {
        switch charset.lowercased() {
        case "utf-8", "utf8":
            .utf8
        case "iso-2022-jp":
            .iso2022JP
        case "shift_jis", "shift-jis", "sjis", "x-sjis":
            .shiftJIS
        case "euc-jp", "eucjp":
            .japaneseEUC
        case "iso-8859-1", "latin1", "latin-1":
            .isoLatin1
        case "us-ascii", "ascii":
            .ascii
        case "big5":
            String.Encoding(rawValue: CFStringConvertEncodingToNSStringEncoding(CFStringEncoding(CFStringEncodings.big5.rawValue)))
        case "gb2312", "gbk":
            String.Encoding(rawValue: CFStringConvertEncodingToNSStringEncoding(CFStringEncoding(CFStringEncodings.GB_18030_2000.rawValue)))
        case "euc-kr":
            String.Encoding(rawValue: CFStringConvertEncodingToNSStringEncoding(CFStringEncoding(CFStringEncodings.EUC_KR.rawValue)))
        default:
            .utf8
        }
    }

    private func removeCRLF(_ value: String) -> String {
        value.replacingOccurrences(of: "\r", with: "").replacingOccurrences(of: "\n", with: " ")
    }
}
