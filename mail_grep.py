from email import policy
from email.header import decode_header
from email.message import Message
from pathlib import Path
import argparse
import csv
import email
import re


class MailFileFinder:
    """emlxファイル探索クラス"""

    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def find(self) -> list[Path]:
        return list(self._root_dir.rglob("*.emlx"))

    # 4メソッドにしたい場合は将来 filter/include_dir 追加も可


class MailTextDecoder:
    """メール1通からデコード済み行群を取り出す"""

    def extract_header_lines(self, msg) -> list[str]:
        return [
            f"{k}: {self._decode_header(msg[k])}"
            for k in ("Subject", "From", "To", "Date")
            if msg[k]
        ]

    def extract_body_lines(self, msg) -> list[str]:
        result = []
        for part in self._iter_text_parts(msg):
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                try:
                    text = payload.decode(charset, errors="replace")
                except Exception:
                    text = payload.decode("utf-8", errors="replace")
                result.extend(text.splitlines())
        return result

    def parse_message(self, raw_bytes: bytes) -> Message:
        raw_str = self._decode_bytes(raw_bytes)
        header_str = self._extract_headers_section(raw_str)
        return email.message_from_string(header_str, policy=policy.default)

    def _decode_header(self, value) -> str:
        if value is None:
            return ""
        parts = decode_header(value)
        out = []
        for text, enc in parts:
            if isinstance(text, bytes):
                if enc:
                    try:
                        return_text = text.decode(enc, errors="strict")
                        out.append(return_text)
                        continue
                    except Exception as e:
                        print(
                            f"[RECOVERY] decode_header: {enc=} decode failed ({e}), trying utf-8 ..."
                        )
                # utf-8リカバリ
                try:
                    return_text = text.decode("utf-8", errors="strict")
                    print(
                        f"[RECOVERY] decode_header: used utf-8 fallback for {repr(text[:32])}"
                    )
                    out.append(return_text)
                    continue
                except Exception:
                    pass
                # latin1リカバリ
                try:
                    return_text = text.decode("latin1", errors="replace")
                    print(
                        f"[RECOVERY] decode_header: used latin1 fallback for {repr(text[:32])}"
                    )
                    out.append(return_text)
                    continue
                except Exception:
                    pass
                # 最後の手段：バイトrepr
                print(
                    f"[RECOVERY] decode_header: could not decode {repr(text[:32])}, outputting raw bytes."
                )
                out.append(repr(text))
            else:
                out.append(str(text))
        return "".join(out)

    def _iter_text_parts(self, msg):
        if msg.is_multipart():
            return (p for p in msg.walk() if p.get_content_type().startswith("text/"))
        return (msg,)

    def _decode_bytes(self, raw: bytes) -> str:
        return raw.decode("utf-8", errors="ignore")

    def _extract_headers_section(self, raw_str: str) -> str:
        lines = []
        for line in raw_str.splitlines():
            if line.strip() == "":
                break
            lines.append(line)
        return "\n".join(lines)


class MailPatternMatcher:
    """パターン設定＆1通からマッチ判定・行抽出"""

    def __init__(self, pattern: str, flags=0):
        self._pattern = re.compile(pattern, flags)

    def match_mail(
        self, msg: Message, decoder: MailTextDecoder
    ) -> tuple[str, str, str, str, str] | None:
        header_lines = decoder.extract_header_lines(msg)
        body_lines = decoder.extract_body_lines(msg)
        all_lines = header_lines + body_lines
        for line in all_lines:
            if self._pattern.search(line):
                return (
                    decoder._decode_header(msg["Subject"]),
                    decoder._decode_header(msg["From"]),
                    decoder._decode_header(msg["To"]),
                    decoder._decode_header(msg["Date"]),
                    line.strip(),
                )
        return None


class MailCsvExporter:
    """全体制御＆CSV出力"""

    def __init__(
        self,
        finder: MailFileFinder,
        decoder: MailTextDecoder,
        matcher: MailPatternMatcher,
        output_path: Path,
    ):
        self._finder = finder
        self._decoder = decoder
        self._matcher = matcher
        self._output_path = output_path

    def process_all(self):
        rows = []
        try:
            for path in self._finder.find():
                raw = self._read_emlx(path)
                msg = self._decoder.parse_message(raw)
                result = self._matcher.match_mail(msg, self._decoder)
                if result:
                    print("✓", result[0], "←", result[4])
                    rows.append(result)
        except KeyboardInterrupt:
            print("\n[INFO] 中断されました。ここまでの結果を保存します。")
        finally:
            self._write_csv(rows)
            print(f"[INFO] {len(rows)}件を {self._output_path} に保存しました。")

    def _read_emlx(self, path: Path) -> bytes:
        raw = path.read_bytes()
        return raw[raw.find(b"\n") + 1 :] if raw[0:1].isdigit() else raw

    def _write_csv(self, rows: list[tuple[str, str, str, str, str]]):
        with open(self._output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Subject", "From", "To", "Date", "Matched Line"])
            writer.writerows(rows)


def egrep_to_python_regex(pattern: str) -> str:
    posix_map = {
        r"\[[:digit:]\]": r"\d",
        r"\[[:space:]\]": r"\s",
        r"\[[:alnum:]\]": r"[A-Za-z0-9]",
        r"\[[:alpha:]\]": r"[A-Za-z]",
        r"\[[:lower:]\]": r"[a-z]",
        r"\[[:upper:]\]": r"[A-Z]",
        r"\[[:punct:]\]": r'[!"#$%&\'()*+,\-./:;<=>?@[\\\]^_`{|}~]',
        r"$begin:math:display$[:blank:]$end:math:display$": r"[ \t]",
        r"$begin:math:display$[:xdigit:]$end:math:display$": r"[A-Fa-f0-9]",
        r"$begin:math:display$[:cntrl:]$end:math:display$": r"[\x00-\x1F\x7F]",
        r"$begin:math:display$[:print:]$end:math:display$": r"[ -~]",
        r"$begin:math:display$[:graph:]$end:math:display$": r"[!-~]",
    }
    for k, v in posix_map.items():
        pattern = re.sub(k, lambda m, v=v: v, pattern)
    return pattern


def create_parser():
    parser = argparse.ArgumentParser(
        description="egrep風にemlxメールをgrepし、CSVに出力するツール"
    )
    parser.add_argument(
        "pattern", metavar="PATTERN", help="検索したい正規表現（egrep互換）"
    )
    parser.add_argument(
        "-i", "--ignore-case", action="store_true", help="大文字・小文字を無視する"
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("output_mail_summary.csv"),
        help="出力CSVファイル名（デフォルト: output_mail_summary.csv）",
    )
    parser.add_argument(
        "-s",
        "--source",
        type=Path,
        default=Path.home() / "Library" / "Mail" / "V10",
        help="emlxファイルの格納ディレクトリ",
    )
    return parser


if __name__ == "__main__":
    parser = create_parser()
    args = parser.parse_args()
    pattern = egrep_to_python_regex(args.pattern)
    flags = re.IGNORECASE if args.ignore_case else 0

    finder = MailFileFinder(args.source)
    decoder = MailTextDecoder()
    matcher = MailPatternMatcher(pattern, flags)
    exporter = MailCsvExporter(finder, decoder, matcher, args.output)
    exporter.process_all()
