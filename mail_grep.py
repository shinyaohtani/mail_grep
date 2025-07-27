from bs4 import BeautifulSoup
from email import policy
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from pathlib import Path
import argparse
import csv
import email
import logging
import os
import re
import sys
import traceback


def parse_date(value: str) -> str:
    try:
        dt = parsedate_to_datetime(value)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return value  # パース不可はそのまま


def remove_crlf(value: str) -> str:
    if not value:
        return ""
    return value.replace("\r", "").replace("\n", " ")


class MailFileFinder:
    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def find(self) -> list[Path]:
        return list(self._root_dir.rglob("*.emlx"))


class MailTextDecoder:
    def get_message_id(self, msg) -> str:
        try:
            v = msg.get("Message-ID")
            v = self._normalize_header_value(v)
            return self._decode_header(v) if v else ""
        except Exception:
            return ""

    def get_date(self, msg) -> str:
        try:
            v = msg.get("Date")
            v = self._normalize_header_value(v)
            if not v:
                return ""
            dt = parsedate_to_datetime(v)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            try:
                v = msg.get("Date")
                v = self._normalize_header_value(v)
                return self._decode_header(v or "")
            except Exception:
                return ""

    def extract_header_lines(self, msg) -> list[str]:
        result = []
        for k in ("Subject", "From", "To", "Date"):
            try:
                v = msg.get(k)
                v = self._normalize_header_value(v)
                if v is None:
                    continue
                decoded = self._decode_header(v)
                decoded = decoded.replace("\r", "").replace("\n", " ")
                result.append(f"{k}: {decoded}")
            except Exception as e:
                logging.warning(f"[MailGrep] Could not parse {k}: {e}")
                result.append(f"{k}: [INVALID HEADER]")
        return result

    def extract_body_lines(self, msg) -> list[tuple[str, str]]:
        """
        本文パート（text/plain, text/html）を(行,パート種別)で返す
        """
        result = []
        for part in self._iter_text_parts(msg):
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                text = payload.decode(charset, errors="replace")
            except Exception:
                text = payload.decode("utf-8", errors="replace")

            content_type = part.get_content_type()
            if content_type == "text/html":
                # HTMLタグ除去
                soup = BeautifulSoup(text, "html.parser")
                text_content = soup.get_text(separator="\n")
                parttype = "html"
            else:
                text_content = text
                parttype = "plain"
            # 各行ごとにtupleで返す
            for line in text_content.splitlines():
                result.append((line, parttype))
        return result

    def parse_message(self, raw_bytes: bytes) -> Message:
        raw_str = self._decode_bytes(raw_bytes)
        header_str = self._extract_headers_section(raw_str)
        header_str = self._sanitize_header_section(header_str)
        safe_header_str = "\n".join(
            [remove_crlf(line) for line in header_str.splitlines()]
        )
        try:
            return email.message_from_string(safe_header_str, policy=policy.default)
        except Exception:
            return email.message_from_string("", policy=policy.default)

    def _sanitize_header_section(self, header_str: str) -> str:
        sanitized = []
        prev_colon = False
        for line in header_str.splitlines():
            if ":" in line:
                key, val = line.split(":", 1)
                val = val.replace("\r", " ").replace("\n", " ")
                sanitized.append(f"{key}:{val}")
                prev_colon = True
            else:
                if prev_colon:
                    line = line.replace("\r", " ").replace("\n", " ")
                sanitized.append(line)
                prev_colon = False
        return "\n".join(sanitized)

    def _normalize_header_value(self, v):
        if v is None:
            return ""
        if hasattr(v, "addresses"):
            try:
                return str(v)
            except Exception:
                return ""
        if isinstance(v, bytes):
            try:
                return v.decode("utf-8", errors="replace")
            except Exception:
                return ""
        if not isinstance(v, str):
            return str(v)
        return v

    def _decode_header(self, value) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            value = remove_crlf(value)
        try:
            parts = decode_header(value)
        except Exception:
            return ""
        out = []
        for text, enc in parts:
            if isinstance(text, bytes):
                status: str = "noErr"
                if enc:
                    try:
                        return_text = text.decode(enc, errors="strict")
                        out.append(return_text)
                        continue
                    except Exception as e:
                        status = f"decode_header: {enc=} decode failed ({e}), trying utf-8 ..."
                out.append(self._decode_header_fallback(text, status))
            else:
                out.append(str(text))
        return "".join(out)

    def _iter_text_parts(self, msg):
        if hasattr(msg, "is_multipart") and msg.is_multipart():
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

    def _decode_header_fallback(self, text: bytes, status: str) -> str:
        try:
            return text.decode("utf-8", errors="strict")
        except Exception:
            pass
        try:
            return text.decode("latin1", errors="replace")
        except Exception:
            pass
        logging.warning(f"[MailGrep] {status}")
        logging.warning(
            f"[MailGrep] decode_header: could not decode {repr(text[:40])}, outputting raw bytes."
        )
        return repr(text)


# --- パターンマッチ ---
class MailPatternMatcher:
    def __init__(self, pattern: str, flags=0):
        self._pattern = re.compile(pattern, flags)

    def match_mail(self, msg: Message, decoder: MailTextDecoder) -> list[tuple]:
        header_lines = decoder.extract_header_lines(msg)
        body_lines = decoder.extract_body_lines(msg)
        matches = []
        # ヘッダーもbody同様にparttypeを付与（"header"固定）
        for line in header_lines:
            if self._pattern.search(line):
                matches.append(("header", line))
        # 本文: (line, parttype)
        for line, parttype in body_lines:
            if self._pattern.search(line):
                matches.append((parttype, line))
        return matches


# --- CSV出力 ---
class MailCsvExporter:
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
        mail_id = 1
        try:
            for path in self._finder.find():
                try:
                    raw = self._read_emlx(path)
                    msg = self._decoder.parse_message(raw)
                    message_id = self._decoder.get_message_id(msg)
                    date_str = self._decoder.get_date(msg)
                    subj = self._decoder._decode_header(
                        self._decoder._normalize_header_value(msg.get("Subject"))
                    )
                    from_ = self._decoder._decode_header(
                        self._decoder._normalize_header_value(msg.get("From"))
                    )
                    to_ = self._decoder._decode_header(
                        self._decoder._normalize_header_value(msg.get("To"))
                    )

                    hit_count = 0
                    for parttype, line in self._matcher.match_mail(msg, self._decoder):
                        hit_count += 1
                        print(f"✓ [{parttype}] {subj} ← {line.strip()}")
                        rows.append(
                            [
                                mail_id,
                                message_id,
                                hit_count,
                                subj,
                                from_,
                                to_,
                                date_str,
                                parttype,
                                line.strip(),
                            ]
                        )
                    mail_id += 1
                except Exception as e:
                    logging.warning(f"[MailGrep] Skipped {path}: {e}")
        except KeyboardInterrupt:
            print()  # 改行を入れる
            logging.warning("[MailGrep]中断されました。ここまでの結果を保存します。")
        finally:
            self._write_csv(rows)
            logging.info(
                f"[MailGrep] {len(rows)}件を {self._output_path} に保存しました。"
            )

    def _read_emlx(self, path: Path) -> bytes:
        raw = path.read_bytes()
        return raw[raw.find(b"\n") + 1 :] if raw[0:1].isdigit() else raw

    def _sanitize_csv_field(self, value) -> str:
        if value is None:
            return ""
        s = str(value)
        return s.replace("\r", "").replace("\n", "⏎")

    def _write_csv(self, rows):
        with open(self._output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "mail_id",
                    "message_id",
                    "hit_id",
                    "Subject",
                    "From",
                    "To",
                    "Date",
                    "Matched Part",
                    "Matched Line",
                ]
            )
            for row in rows:
                writer.writerow([self._sanitize_csv_field(v) for v in row])


# --- egrep to Python正規表現（必要なら改良して下さい） ---
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


# --- 引数パーサ ---
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


# --- メイン処理 ---
def main():
    parser = create_parser()
    args = parser.parse_args()
    pattern = egrep_to_python_regex(args.pattern)
    flags = re.IGNORECASE if args.ignore_case else 0

    finder = MailFileFinder(args.source)
    decoder = MailTextDecoder()
    matcher = MailPatternMatcher(pattern, flags)
    exporter = MailCsvExporter(finder, decoder, matcher, args.output)
    exporter.process_all()


if __name__ == "__main__":
    log_level = (
        logging.DEBUG if os.getenv("MAIL_GREP_DEBUG", "0") == "1" else logging.INFO
    )
    logging.basicConfig(
        level=log_level, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    try:
        main()
    except Exception as e:
        logging.error(f"[MailGrep] Unhandled exception: {e}")
        logging.error("Unhandled exception at top-level:\n" + traceback.format_exc())
        print("==== FATAL TRACEBACK ====")
        traceback.print_exc(file=sys.stdout)
        exit(1)
