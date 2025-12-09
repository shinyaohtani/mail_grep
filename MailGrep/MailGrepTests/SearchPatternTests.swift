@testable import MailGrep
import XCTest

/// egrep互換の正規表現テスト (80テスト)
/// 各パターンが正しくマッチし、正しくマッチしないことを確認
final class SearchPatternTests: XCTestCase {
    // MARK: - 1. リテラル文字列マッチ (4テスト)

    /// 1. 単純な文字列マッチ
    func testLiteralSimpleMatch() throws {
        let sp = try SearchPattern(pattern: "hello", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello world"))
        XCTAssertFalse(sp.matches("goodbye world"))
    }

    /// 2. 複数単語のリテラルマッチ
    func testLiteralMultiWordMatch() throws {
        let sp = try SearchPattern(pattern: "hello world", ignoreCase: false)
        XCTAssertTrue(sp.matches("say hello world today"))
        XCTAssertFalse(sp.matches("helloworld"))
    }

    /// 3. 日本語リテラルマッチ
    func testLiteralJapaneseMatch() throws {
        let sp = try SearchPattern(pattern: "こんにちは", ignoreCase: false)
        XCTAssertTrue(sp.matches("今日はこんにちは世界"))
        XCTAssertFalse(sp.matches("さようなら世界"))
    }

    /// 4. 数字リテラルマッチ
    func testLiteralNumberMatch() throws {
        let sp = try SearchPattern(pattern: "12345", ignoreCase: false)
        XCTAssertTrue(sp.matches("order 12345 confirmed"))
        XCTAssertFalse(sp.matches("order 54321 confirmed"))
    }

    // MARK: - 2. ドット（任意の1文字）(4テスト)

    /// 5. ドット基本マッチ
    func testDotBasicMatch() throws {
        let sp = try SearchPattern(pattern: "h.llo", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertTrue(sp.matches("hallo"))
        XCTAssertFalse(sp.matches("hllo"))
    }

    /// 6. 複数ドットマッチ
    func testDotMultipleMatch() throws {
        let sp = try SearchPattern(pattern: "t..t", ignoreCase: false)
        XCTAssertTrue(sp.matches("test"))
        XCTAssertTrue(sp.matches("that"))
        XCTAssertFalse(sp.matches("tt"))
    }

    /// 7. ドット先頭マッチ
    func testDotAtStartMatch() throws {
        let sp = try SearchPattern(pattern: ".est", ignoreCase: false)
        XCTAssertTrue(sp.matches("test"))
        XCTAssertTrue(sp.matches("best"))
        XCTAssertFalse(sp.matches("est"))
    }

    /// 8. ドット末尾マッチ
    func testDotAtEndMatch() throws {
        let sp = try SearchPattern(pattern: "tes.", ignoreCase: false)
        XCTAssertTrue(sp.matches("test"))
        XCTAssertTrue(sp.matches("tess"))
        XCTAssertFalse(sp.matches("tes"))
    }

    // MARK: - 3. アスタリスク量指定子（0回以上）(4テスト)

    /// 9. アスタリスク基本マッチ
    func testStarBasicMatch() throws {
        let sp = try SearchPattern(pattern: "hel*o", ignoreCase: false)
        XCTAssertTrue(sp.matches("heo"))
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertTrue(sp.matches("helllo"))
    }

    /// 10. アスタリスク0回マッチ
    func testStarZeroMatch() throws {
        let sp = try SearchPattern(pattern: "ab*c", ignoreCase: false)
        XCTAssertTrue(sp.matches("ac"))
        XCTAssertFalse(sp.matches("ad"))
    }

    /// 11. ドット+アスタリスク（任意文字列）
    func testDotStarMatch() throws {
        let sp = try SearchPattern(pattern: "start.*end", ignoreCase: false)
        XCTAssertTrue(sp.matches("start middle end"))
        XCTAssertTrue(sp.matches("startend"))
        XCTAssertFalse(sp.matches("start only"))
    }

    /// 12. アスタリスク貪欲マッチ
    func testStarGreedyMatch() throws {
        let sp = try SearchPattern(pattern: "a.*b", ignoreCase: false)
        XCTAssertTrue(sp.matches("aXXXbYYYb"))
        XCTAssertFalse(sp.matches("aXXX"))
    }

    // MARK: - 4. プラス量指定子（1回以上）(4テスト)

    /// 13. プラス基本マッチ
    func testPlusBasicMatch() throws {
        let sp = try SearchPattern(pattern: "hel+o", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertTrue(sp.matches("helllo"))
        XCTAssertFalse(sp.matches("heo"))
    }

    /// 14. プラス1回マッチ
    func testPlusOneMatch() throws {
        let sp = try SearchPattern(pattern: "ab+c", ignoreCase: false)
        XCTAssertTrue(sp.matches("abc"))
        XCTAssertFalse(sp.matches("ac"))
    }

    /// 15. ドット+プラス（1文字以上）
    func testDotPlusMatch() throws {
        let sp = try SearchPattern(pattern: "a.+b", ignoreCase: false)
        XCTAssertTrue(sp.matches("aXb"))
        XCTAssertTrue(sp.matches("aXXXb"))
        XCTAssertFalse(sp.matches("ab"))
    }

    /// 16. プラス複数適用
    func testPlusMultipleMatch() throws {
        let sp = try SearchPattern(pattern: "a+b+c+", ignoreCase: false)
        XCTAssertTrue(sp.matches("abc"))
        XCTAssertTrue(sp.matches("aaabbbccc"))
        XCTAssertFalse(sp.matches("ac"))
    }

    // MARK: - 5. クエスチョン量指定子（0回または1回）(4テスト)

    /// 17. クエスチョン基本マッチ
    func testQuestionBasicMatch() throws {
        let sp = try SearchPattern(pattern: "colou?r", ignoreCase: false)
        XCTAssertTrue(sp.matches("color"))
        XCTAssertTrue(sp.matches("colour"))
        XCTAssertFalse(sp.matches("colouur"))
    }

    /// 18. クエスチョン0回マッチ
    func testQuestionZeroMatch() throws {
        let sp = try SearchPattern(pattern: "ab?c", ignoreCase: false)
        XCTAssertTrue(sp.matches("ac"))
        XCTAssertTrue(sp.matches("abc"))
        XCTAssertFalse(sp.matches("abbc"))
    }

    /// 19. クエスチョン複数適用
    func testQuestionMultipleMatch() throws {
        let sp = try SearchPattern(pattern: "^a?b?c$", ignoreCase: false) // アンカーで厳密マッチ
        XCTAssertTrue(sp.matches("c"))
        XCTAssertTrue(sp.matches("abc"))
        XCTAssertFalse(sp.matches("aabc"))
    }

    /// 20. オプショナルグループ
    func testQuestionGroupMatch() throws {
        let sp = try SearchPattern(pattern: "^(un)?happy$", ignoreCase: false) // アンカーで厳密マッチ
        XCTAssertTrue(sp.matches("happy"))
        XCTAssertTrue(sp.matches("unhappy"))
        XCTAssertFalse(sp.matches("ununhappy"))
    }

    // MARK: - 6. 回数指定 {n}, {n,}, {n,m} (4テスト)

    /// 21. 固定回数 {n}
    func testBraceExactMatch() throws {
        let sp = try SearchPattern(pattern: "a{3}", ignoreCase: false)
        XCTAssertTrue(sp.matches("aaa"))
        XCTAssertTrue(sp.matches("baaab"))
        XCTAssertFalse(sp.matches("aa"))
    }

    /// 22. 最小回数 {n,}
    func testBraceMinMatch() throws {
        let sp = try SearchPattern(pattern: "a{2,}", ignoreCase: false)
        XCTAssertTrue(sp.matches("aa"))
        XCTAssertTrue(sp.matches("aaaa"))
        XCTAssertFalse(sp.matches("a"))
    }

    /// 23. 範囲指定 {n,m}
    func testBraceRangeMatch() throws {
        let sp = try SearchPattern(pattern: "a{2,4}", ignoreCase: false)
        XCTAssertTrue(sp.matches("aa"))
        XCTAssertTrue(sp.matches("aaaa"))
        XCTAssertFalse(sp.matches("a"))
    }

    /// 24. 複合回数指定
    func testBraceComplexMatch() throws {
        let sp = try SearchPattern(pattern: "[0-9]{3}-[0-9]{4}", ignoreCase: false)
        XCTAssertTrue(sp.matches("123-4567"))
        XCTAssertFalse(sp.matches("12-345"))
    }

    // MARK: - 7. 文字クラス [abc] (4テスト)

    /// 25. 文字クラス基本
    func testCharClassBasicMatch() throws {
        let sp = try SearchPattern(pattern: "[aeiou]", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertFalse(sp.matches("rhythm"))
    }

    /// 26. 文字クラス範囲
    func testCharClassRangeMatch() throws {
        let sp = try SearchPattern(pattern: "[a-z]", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertFalse(sp.matches("12345"))
    }

    /// 27. 数字クラス範囲
    func testCharClassDigitRangeMatch() throws {
        let sp = try SearchPattern(pattern: "[0-9]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("abc123def"))
        XCTAssertFalse(sp.matches("abcdef"))
    }

    /// 28. 複合文字クラス
    func testCharClassComplexMatch() throws {
        let sp = try SearchPattern(pattern: "[a-zA-Z0-9_]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("user_123"))
        XCTAssertFalse(sp.matches("@#$%"))
    }

    // MARK: - 8. 否定文字クラス [^abc] (4テスト)

    /// 29. 否定文字クラス基本
    func testNegCharClassBasicMatch() throws {
        let sp = try SearchPattern(pattern: "[^aeiou]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("rhythm"))
        XCTAssertFalse(sp.matches("aaa"))
    }

    /// 30. 否定数字クラス
    func testNegCharClassDigitMatch() throws {
        let sp = try SearchPattern(pattern: "[^0-9]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("abc"))
        XCTAssertFalse(sp.matches("123"))
    }

    /// 31. 否定範囲クラス
    func testNegCharClassRangeMatch() throws {
        let sp = try SearchPattern(pattern: "[^a-z]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("ABC123"))
        XCTAssertFalse(sp.matches("abc"))
    }

    /// 32. 否定複合クラス
    func testNegCharClassComplexMatch() throws {
        let sp = try SearchPattern(pattern: "[^\\s]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertFalse(sp.matches("   "))
    }

    // MARK: - 9. アンカー ^ と $ (4テスト)

    /// 33. 行頭アンカー
    func testAnchorStartMatch() throws {
        let sp = try SearchPattern(pattern: "^hello", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello world"))
        XCTAssertFalse(sp.matches("say hello"))
    }

    /// 34. 行末アンカー
    func testAnchorEndMatch() throws {
        let sp = try SearchPattern(pattern: "world$", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello world"))
        XCTAssertFalse(sp.matches("world hello"))
    }

    /// 35. 両端アンカー
    func testAnchorBothMatch() throws {
        let sp = try SearchPattern(pattern: "^hello$", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertFalse(sp.matches("hello world"))
    }

    /// 36. アンカー+パターン
    func testAnchorPatternMatch() throws {
        let sp = try SearchPattern(pattern: "^[A-Z]", ignoreCase: false)
        XCTAssertTrue(sp.matches("Hello"))
        XCTAssertFalse(sp.matches("hello"))
    }

    // MARK: - 10. 選択（オルタネーション）| (4テスト)

    /// 37. 選択基本
    func testAlternationBasicMatch() throws {
        let sp = try SearchPattern(pattern: "cat|dog", ignoreCase: false)
        XCTAssertTrue(sp.matches("I have a cat"))
        XCTAssertTrue(sp.matches("I have a dog"))
        XCTAssertFalse(sp.matches("I have a bird"))
    }

    /// 38. 複数選択
    func testAlternationMultipleMatch() throws {
        let sp = try SearchPattern(pattern: "red|green|blue", ignoreCase: false)
        XCTAssertTrue(sp.matches("color is red"))
        XCTAssertTrue(sp.matches("color is blue"))
        XCTAssertFalse(sp.matches("color is yellow"))
    }

    /// 39. グループ内選択
    func testAlternationGroupMatch() throws {
        let sp = try SearchPattern(pattern: "(Mr|Mrs|Ms)\\. Smith", ignoreCase: false)
        XCTAssertTrue(sp.matches("Mr. Smith"))
        XCTAssertTrue(sp.matches("Mrs. Smith"))
        XCTAssertFalse(sp.matches("Dr. Smith"))
    }

    /// 40. 日本語選択
    func testAlternationJapaneseMatch() throws {
        let sp = try SearchPattern(pattern: "東京|大阪|名古屋", ignoreCase: false)
        XCTAssertTrue(sp.matches("出張先: 東京"))
        XCTAssertTrue(sp.matches("出張先: 大阪"))
        XCTAssertFalse(sp.matches("出張先: 福岡"))
    }

    // MARK: - 11. グルーピング () (4テスト)

    /// 41. グループ基本
    func testGroupBasicMatch() throws {
        let sp = try SearchPattern(pattern: "(ab)+", ignoreCase: false)
        XCTAssertTrue(sp.matches("abab"))
        XCTAssertTrue(sp.matches("ababab"))
        XCTAssertFalse(sp.matches("acac")) // "ab"のシーケンスがない
    }

    /// 42. ネストグループ
    func testGroupNestedMatch() throws {
        let sp = try SearchPattern(pattern: "((a|b)+c)+", ignoreCase: false)
        XCTAssertTrue(sp.matches("abc"))
        XCTAssertTrue(sp.matches("aabcbbc"))
        XCTAssertFalse(sp.matches("ab"))
    }

    /// 43. グループ+量指定子
    func testGroupQuantifierMatch() throws {
        let sp = try SearchPattern(pattern: "(ha){2,}", ignoreCase: false)
        XCTAssertTrue(sp.matches("haha"))
        XCTAssertTrue(sp.matches("hahaha"))
        XCTAssertFalse(sp.matches("ha"))
    }

    /// 44. 非キャプチャグループ (egrep互換)
    func testGroupNonCapturingMatch() throws {
        let sp = try SearchPattern(pattern: "(?:foo|bar)baz", ignoreCase: false)
        XCTAssertTrue(sp.matches("foobaz"))
        XCTAssertTrue(sp.matches("barbaz"))
        XCTAssertFalse(sp.matches("bazbaz"))
    }

    // MARK: - 12. エスケープシーケンス (4テスト)

    /// 45. \\d 数字
    func testEscapeDigitMatch() throws {
        let sp = try SearchPattern(pattern: "\\d+", ignoreCase: false)
        XCTAssertTrue(sp.matches("abc123def"))
        XCTAssertFalse(sp.matches("abcdef"))
    }

    /// 46. \\w 単語文字
    func testEscapeWordMatch() throws {
        let sp = try SearchPattern(pattern: "\\w+", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello123"))
        XCTAssertFalse(sp.matches("@#$%"))
    }

    /// 47. \\s 空白
    func testEscapeSpaceMatch() throws {
        let sp = try SearchPattern(pattern: "hello\\s+world", ignoreCase: false)
        XCTAssertTrue(sp.matches("hello world"))
        XCTAssertTrue(sp.matches("hello   world"))
        XCTAssertFalse(sp.matches("helloworld"))
    }

    /// 48. \\b 単語境界
    func testEscapeBoundaryMatch() throws {
        let sp = try SearchPattern(pattern: "\\bword\\b", ignoreCase: false)
        XCTAssertTrue(sp.matches("a word here"))
        XCTAssertFalse(sp.matches("keyword"))
    }

    // MARK: - 13. 特殊文字エスケープ (4テスト)

    /// 49. ドットエスケープ
    func testEscapeDotMatch() throws {
        let sp = try SearchPattern(pattern: "example\\.com", ignoreCase: false)
        XCTAssertTrue(sp.matches("www.example.com"))
        XCTAssertFalse(sp.matches("www.exampleXcom"))
    }

    /// 50. アスタリスクエスケープ
    func testEscapeStarMatch() throws {
        let sp = try SearchPattern(pattern: "a\\*b", ignoreCase: false)
        XCTAssertTrue(sp.matches("a*b"))
        XCTAssertFalse(sp.matches("ab"))
    }

    /// 51. プラスエスケープ
    func testEscapePlusMatch() throws {
        let sp = try SearchPattern(pattern: "a\\+b", ignoreCase: false)
        XCTAssertTrue(sp.matches("a+b"))
        XCTAssertFalse(sp.matches("ab"))
    }

    /// 52. 括弧エスケープ
    func testEscapeParenMatch() throws {
        let sp = try SearchPattern(pattern: "\\(test\\)", ignoreCase: false)
        XCTAssertTrue(sp.matches("(test)"))
        XCTAssertFalse(sp.matches("test"))
    }

    // MARK: - 14. 大文字小文字 (4テスト)

    /// 53. 大文字小文字区別あり
    func testCaseSensitiveMatch() throws {
        let sp = try SearchPattern(pattern: "Hello", ignoreCase: false)
        XCTAssertTrue(sp.matches("Hello"))
        XCTAssertFalse(sp.matches("hello"))
        XCTAssertFalse(sp.matches("HELLO"))
    }

    /// 54. 大文字小文字区別なし
    func testCaseInsensitiveMatch() throws {
        let sp = try SearchPattern(pattern: "hello", ignoreCase: true)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertTrue(sp.matches("Hello"))
        XCTAssertTrue(sp.matches("HELLO"))
    }

    /// 55. 混合ケースパターン
    func testCaseMixedPatternMatch() throws {
        let sp = try SearchPattern(pattern: "HeLLo", ignoreCase: true)
        XCTAssertTrue(sp.matches("hello"))
        XCTAssertTrue(sp.matches("HELLO"))
    }

    /// 56. 日本語+英語混合
    func testCaseJapaneseEnglishMatch() throws {
        let sp = try SearchPattern(pattern: "テストTest", ignoreCase: true)
        XCTAssertTrue(sp.matches("テストtest"))
        XCTAssertTrue(sp.matches("テストTEST"))
    }

    // MARK: - 15. メールアドレスパターン (4テスト)

    /// 57. 基本メールパターン
    func testEmailBasicMatch() throws {
        let sp = try SearchPattern(pattern: "[\\w.+-]+@[\\w-]+\\.[\\w.-]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("test@example.com"))
        XCTAssertFalse(sp.matches("not-an-email"))
    }

    /// 58. サブドメインメール
    func testEmailSubdomainMatch() throws {
        let sp = try SearchPattern(pattern: "[\\w.+-]+@[\\w.-]+\\.[a-z]{2,}", ignoreCase: true)
        XCTAssertTrue(sp.matches("user@mail.example.co.jp"))
        XCTAssertFalse(sp.matches("user@localhost"))
    }

    /// 59. プラス記号付きメール
    func testEmailPlusMatch() throws {
        let sp = try SearchPattern(pattern: "[\\w.+-]+@[\\w-]+\\.[\\w]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("user+tag@example.com"))
        XCTAssertFalse(sp.matches("@example.com"))
    }

    /// 60. メール抽出
    func testEmailExtractMatch() throws {
        let sp = try SearchPattern(pattern: "From:.*<[^>]+@[^>]+>", ignoreCase: true)
        XCTAssertTrue(sp.matches("From: John <john@example.com>"))
        XCTAssertFalse(sp.matches("From: John"))
    }

    // MARK: - 16. URLパターン (4テスト)

    /// 61. 基本URLパターン
    func testUrlBasicMatch() throws {
        let sp = try SearchPattern(pattern: "https?://[\\w.-]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("https://example.com"))
        XCTAssertTrue(sp.matches("http://test.org"))
        XCTAssertFalse(sp.matches("ftp://server"))
    }

    /// 62. パス付きURL
    func testUrlPathMatch() throws {
        let sp = try SearchPattern(pattern: "https?://[\\w.-]+/[\\w./-]*", ignoreCase: false)
        XCTAssertTrue(sp.matches("https://example.com/path/to/page"))
        XCTAssertFalse(sp.matches("https://example.com"))
    }

    /// 63. クエリパラメータ付きURL
    func testUrlQueryMatch() throws {
        let sp = try SearchPattern(pattern: "https?://[^\\s]+\\?[^\\s]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("https://example.com?id=123"))
        XCTAssertFalse(sp.matches("https://example.com"))
    }

    /// 64. 日本語ドメイン
    func testUrlJapaneseDomainMatch() throws {
        let sp = try SearchPattern(pattern: "https?://[^\\s]+", ignoreCase: false)
        XCTAssertTrue(sp.matches("https://日本語.jp"))
        XCTAssertFalse(sp.matches("not a url"))
    }

    // MARK: - 17. 日付パターン (4テスト)

    /// 65. ISO日付形式
    func testDateIsoMatch() throws {
        let sp = try SearchPattern(pattern: "\\d{4}-\\d{2}-\\d{2}", ignoreCase: false)
        XCTAssertTrue(sp.matches("2024-01-15"))
        XCTAssertFalse(sp.matches("01-15-2024"))
    }

    /// 66. 日本式日付形式
    func testDateJapaneseMatch() throws {
        let sp = try SearchPattern(pattern: "\\d{4}年\\d{1,2}月\\d{1,2}日", ignoreCase: false)
        XCTAssertTrue(sp.matches("2024年1月15日"))
        XCTAssertTrue(sp.matches("2024年12月31日"))
        XCTAssertFalse(sp.matches("2024/01/15"))
    }

    /// 67. 時刻形式
    func testTimeMatch() throws {
        let sp = try SearchPattern(pattern: "\\d{2}:\\d{2}(:\\d{2})?", ignoreCase: false)
        XCTAssertTrue(sp.matches("14:30"))
        XCTAssertTrue(sp.matches("14:30:45"))
        XCTAssertFalse(sp.matches("14-30"))
    }

    /// 68. 日時形式
    func testDateTimeMatch() throws {
        let sp = try SearchPattern(pattern: "\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}", ignoreCase: false)
        XCTAssertTrue(sp.matches("2024-01-15T14:30"))
        XCTAssertFalse(sp.matches("2024-01-15 14:30"))
    }

    // MARK: - 18. 電話番号パターン (4テスト)

    /// 69. 日本電話番号（ハイフン区切り）
    func testPhoneJapaneseHyphenMatch() throws {
        let sp = try SearchPattern(pattern: "0\\d{1,4}-\\d{1,4}-\\d{4}", ignoreCase: false)
        XCTAssertTrue(sp.matches("03-1234-5678"))
        XCTAssertTrue(sp.matches("090-1234-5678"))
        XCTAssertFalse(sp.matches("1234-5678"))
    }

    /// 70. 日本電話番号（括弧形式）
    func testPhoneJapaneseParenMatch() throws {
        let sp = try SearchPattern(pattern: "\\(0\\d{1,4}\\)\\d{1,4}-\\d{4}", ignoreCase: false)
        XCTAssertTrue(sp.matches("(03)1234-5678"))
        XCTAssertFalse(sp.matches("03-1234-5678"))
    }

    /// 71. 携帯電話番号
    func testPhoneMobileMatch() throws {
        let sp = try SearchPattern(pattern: "0[789]0-\\d{4}-\\d{4}", ignoreCase: false)
        XCTAssertTrue(sp.matches("090-1234-5678"))
        XCTAssertTrue(sp.matches("080-1234-5678"))
        XCTAssertFalse(sp.matches("03-1234-5678"))
    }

    /// 72. 国際電話番号
    func testPhoneInternationalMatch() throws {
        let sp = try SearchPattern(pattern: "\\+\\d{1,3}-\\d+-\\d+", ignoreCase: false)
        XCTAssertTrue(sp.matches("+81-3-1234-5678"))
        XCTAssertFalse(sp.matches("03-1234-5678"))
    }

    // MARK: - 19. 金額パターン (4テスト)

    /// 73. 円金額
    func testMoneyYenMatch() throws {
        let sp = try SearchPattern(pattern: "[¥￥]?[0-9,]+円", ignoreCase: false)
        XCTAssertTrue(sp.matches("1,000円"))
        XCTAssertTrue(sp.matches("¥10,000円"))
        XCTAssertFalse(sp.matches("$100"))
    }

    /// 74. ドル金額
    func testMoneyDollarMatch() throws {
        let sp = try SearchPattern(pattern: "\\$[0-9,]+(\\.[0-9]{2})?", ignoreCase: false)
        XCTAssertTrue(sp.matches("$1,000"))
        XCTAssertTrue(sp.matches("$1,000.50"))
        XCTAssertFalse(sp.matches("1000円"))
    }

    /// 75. カンマ区切り数値
    func testMoneyCommaMatch() throws {
        let sp = try SearchPattern(pattern: "^[0-9]{1,3}(,[0-9]{3})+$", ignoreCase: false) // カンマ必須+アンカー
        XCTAssertTrue(sp.matches("1,234,567"))
        XCTAssertFalse(sp.matches("1234567"))
    }

    /// 76. 小数点付き金額
    func testMoneyDecimalMatch() throws {
        let sp = try SearchPattern(pattern: "[0-9]+\\.[0-9]{2}", ignoreCase: false)
        XCTAssertTrue(sp.matches("123.45"))
        XCTAssertFalse(sp.matches("123"))
    }

    // MARK: - 20. 非貪欲マッチ (4テスト)

    /// 77. 非貪欲アスタリスク
    func testNonGreedyStarMatch() throws {
        let sp = try SearchPattern(pattern: "<.*?>", ignoreCase: false)
        XCTAssertTrue(sp.matches("<tag>content</tag>"))
    }

    /// 78. 非貪欲プラス
    func testNonGreedyPlusMatch() throws {
        let sp = try SearchPattern(pattern: "<.+?>", ignoreCase: false)
        XCTAssertTrue(sp.matches("<tag>content</tag>"))
    }

    /// 79. 非貪欲クエスチョン
    func testNonGreedyQuestionMatch() throws {
        let sp = try SearchPattern(pattern: "a.??b", ignoreCase: false)
        XCTAssertTrue(sp.matches("ab"))
        XCTAssertTrue(sp.matches("aXb"))
    }

    /// 80. 非貪欲回数指定
    func testNonGreedyBraceMatch() throws {
        let sp = try SearchPattern(pattern: "a{2,4}?", ignoreCase: false)
        XCTAssertTrue(sp.matches("aa"))
        XCTAssertTrue(sp.matches("aaaa"))
    }

    // MARK: - uniqueName tests (既存の6テスト)

    func testUniqueNameSimple() throws {
        let sp = try SearchPattern(pattern: "invoice", ignoreCase: false)
        let name = sp.uniqueName
        XCTAssertTrue(name.hasPrefix("invoice_"))
    }

    func testUniqueNameWithSpaces() throws {
        let sp = try SearchPattern(pattern: "hello world", ignoreCase: false)
        let name = sp.uniqueName
        XCTAssertTrue(name.contains("hello") || name.contains("world"))
    }

    func testUniqueNameSpecialCharsRemoved() throws {
        let sp = try SearchPattern(pattern: "test.*pattern", ignoreCase: false)
        let name = sp.uniqueName
        XCTAssertFalse(name.contains("*"))
    }

    func testUniqueNameTruncatedAt16() throws {
        let sp = try SearchPattern(pattern: "verylongpatternthatshouldbetruncated", ignoreCase: false)
        let name = sp.uniqueName
        let prefix = name.components(separatedBy: "_20").first ?? ""
        XCTAssertLessThanOrEqual(prefix.count, 16)
    }

    func testUniqueNameJapanese() throws {
        let sp = try SearchPattern(pattern: "請求書", ignoreCase: false)
        let name = sp.uniqueName
        XCTAssertTrue(name.contains("請求書"))
    }

    func testUniqueNameEmpty() throws {
        let sp = try SearchPattern(pattern: ".*", ignoreCase: false)
        let name = sp.uniqueName
        XCTAssertTrue(name.hasPrefix("search_"))
    }
}
