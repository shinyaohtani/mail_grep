from bs4 import BeautifulSoup
from datetime import datetime
from email import policy
from email.header import decode_header
from email.message import Message
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import quote
import argparse
import csv
import email
import logging
import os
import re
import sys
import traceback
import unicodedata

from openpyxl import Workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet
from typing import cast


from smart_logging import SmartLogging


class MailFolder:
    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def mail_paths(self) -> list[Path]:
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
        result = []
        for part in self._iter_text_parts(msg):
            payload = part.get_payload(decode=True)
            if not payload:
                continue

            charset = part.get_content_charset() or "utf-8"
            try:
                html = payload.decode(charset, errors="replace")
            except Exception:
                html = payload.decode("utf-8", errors="replace")

            ctype = part.get_content_type()
            if ctype == "text/html":
                # ① 生HTMLをそのまま
                result.append((html, "text/html"))

                # ② BeautifulSoupでパース
                soup = BeautifulSoup(html, "html.parser")
                #    不要タグをまるごと削除
                for tag in soup(["head", "script", "style", "meta", "title", "link"]):
                    tag.decompose()

                # ③ テキストのみ改行区切りで取得
                text_only = soup.get_text(separator="\n", strip=True)
                for line in text_only.splitlines():
                    if line.strip():
                        result.append((line, "text/html_textonly"))

                # ④ さらに全文連結版も追加
                concat = "".join(soup.stripped_strings)
                if concat:
                    result.append((concat, "text/html_concat"))

            else:
                # text/plain
                text = html
                for line in text.splitlines():
                    if line.strip():
                        result.append((line, ctype))

        return result

    def parse_message(self, raw_bytes: bytes) -> Message:
        """
        .emlx の長さ行をスキップしたあと、
        ヘッダー＋本文をそのまま BytesParser に渡す
        """
        # 1) 長さ行があればスキップ
        data = raw_bytes
        if data[:1].isdigit():
            nl = data.find(b"\n")
            if nl != -1:
                data = data[nl + 1 :]

        # 2) バイト列を丸ごとパース（ヘッダーも本文もそのまま）
        parser = BytesParser(policy=policy.default)
        try:
            return parser.parsebytes(data)
        except Exception:
            # 万一エラーが起きたらフォールバック
            return email.message_from_bytes(data, policy=policy.default)

    def get_date_dt(self, msg) -> datetime | None:
        """並べ替え用に Date を datetime で返す（失敗時は None）"""
        try:
            v = msg.get("Date")
            v = self._normalize_header_value(v)
            if not v:
                return None
            return parsedate_to_datetime(v)
        except Exception:
            return None

    def build_mail_link(self, message_id: str) -> str:
        """
        Mail.app で開けるリンクを生成。
        形式: message:%3Cmessage-id%3E  （< と > は URL エンコード）
        """
        if not message_id:
            return ""
        mid = message_id.strip()
        if not mid.startswith("<"):
            mid = f"<{mid}>"
        # 角括弧や@を含めて完全にエンコード（Excel でもクリック可能）
        return f"message:{quote(mid, safe='')}"

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
            value = self.remove_crlf(value)
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

    @staticmethod
    def remove_crlf(value: str) -> str:
        if not value:
            return ""
        return value.replace("\r", "").replace("\n", " ")


# --- パターンマッチ ---
class SearchPattern:
    def __init__(self, egrep_pattern: str, flags=0):
        self._egrep_pattern = egrep_pattern
        python_pattern = self.egrep_to_python_regex(self._egrep_pattern)
        self._pattern = re.compile(python_pattern, flags | re.DOTALL)

    # --- egrep to Python正規表現（必要なら改良して下さい） ---
    @staticmethod
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

    def match_mail(self, msg: Message, decoder: MailTextDecoder) -> list[tuple]:
        matches: list[tuple[str, str]] = []
        # 1) ヘッダー行はこれまでどおり検索
        for line in decoder.extract_header_lines(msg):
            if self._pattern.search(line):
                matches.append(("header", line))

        # 2) 本文行は text/plain と text/html_textonly のみ対象
        for line, parttype in decoder.extract_body_lines(msg):
            if parttype not in ("text/plain", "text/html_textonly"):
                continue
            if self._pattern.search(line):
                matches.append((parttype, line))

        return matches

    def uid(self) -> str:
        """
        検索文字列からデフォルト保存先:
        <cleaned_head16>_<YYYY-MM-DD_HHMMSS>
        - 空白は "_"、正規表現の特殊文字は除去
        - 先頭の非特殊文字のみから最大16文字
        """
        # NFKC正規化 → 空白を "_" に
        s = unicodedata.normalize("NFKC", self._egrep_pattern).replace(" ", "_")
        # アルファベット・数字・アンダースコア以外を除去
        s = re.sub(r"[^\w]", "", s)
        # 連続 "_" の圧縮＆前後の "_" を除去
        s = re.sub(r"_+", "_", s).strip("_")
        # 先頭16文字（空なら "search"）
        head = (s[:16] or "search").lower()
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        return f"{head}_{ts}"


# --- CSV出力 ---
class MailCsvExporter:
    def __init__(
        self,
        finder: MailFolder,
        decoder: MailTextDecoder,
        matcher: SearchPattern,
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
            for path in self._finder.mail_paths():
                try:
                    raw = self._read_emlx(path)
                    msg = self._decoder.parse_message(raw)

                    message_id = self._decoder.get_message_id(msg)
                    date_str = self._decoder.get_date(
                        msg
                    )  # 表示用（YYYY-MM-DD HH:MM:SS or 原文）
                    date_dt = self._decoder.get_date_dt(msg)  # 並べ替え用
                    subj = self._decoder._decode_header(
                        self._decoder._normalize_header_value(msg.get("Subject"))
                    )
                    from_ = self._decoder._decode_header(
                        self._decoder._normalize_header_value(msg.get("From"))
                    )
                    to_ = self._decoder._decode_header(
                        self._decoder._normalize_header_value(msg.get("To"))
                    )
                    link = self._decoder.build_mail_link(message_id)

                    hit_count = 0
                    for parttype, line in self._matcher.match_mail(msg, self._decoder):
                        hit_count += 1
                        max_len = 50
                        preview = line.strip()
                        if len(preview) > max_len:
                            preview = preview[:max_len] + "..."
                        if hit_count == 1:
                            print(f"✓ [{parttype}] hit! ({date_str}) {preview}")
                        else:
                            print(f"✓ [{parttype}] {preview}")
                        # 並べ替え用キーを先頭に持たせておく（書き出し時に除去）
                        rows.append(
                            [
                                date_dt,  # sort key (datetime or None)
                                mail_id,  # mail_id
                                hit_count,  # hit_id
                                message_id,  # message_id
                                link,  # link (Mail.app)
                                date_str,  # Date (表示用)
                                from_,
                                to_,
                                subj,
                                parttype,  # Matched Part
                                line.strip(),  # Matched Line
                            ]
                        )
                    mail_id += 1
                except Exception as e:
                    logging.warning(f"[MailGrep] Skipped {path}: {e}")
        except KeyboardInterrupt:
            print()
            logging.warning("[MailGrep]中断されました。ここまでの結果を保存します。")
        finally:
            # 1) Date 降順でソート（None は末尾）
            def sort_key(row):
                dt = row[0]
                return (dt is None, -(dt.timestamp() if dt else 0))

            rows.sort(key=sort_key)

            # 2) メール件数を計算（hit_id == 1 の行のみカウント）
            mail_count = sum(1 for row in rows if row[2] == 1)

            # 拡張子が何であろうと2種類とも保存する
            self._write_xlsx(rows)
            self._write_csv(rows)
            logging.info(
                f"[MailGrep] {mail_count}件を {self._output_path} に保存しました。"
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
        # 4) BOM 付き UTF-8 で保存（Excel 配慮）
        csv_path = self._output_path.with_suffix(".csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            # 3) 指定順のカラムでヘッダ行を書き込み
            writer.writerow(
                [
                    "mail_id",
                    "hit_id",
                    "message_id",
                    "link",
                    "Date",
                    "From",
                    "To",
                    "Subject",
                    "Matched Part",
                    "Matched Line",
                ]
            )
            for row in rows:
                # 先頭の sort key を捨て、指定順に並べ替えて書き出し
                (
                    _,
                    mail_id,
                    hit_id,
                    message_id,
                    link,
                    date_str,
                    from_,
                    to_,
                    subj,
                    parttype,
                    matched,
                ) = row
                excel_link = f'=HYPERLINK("{link}","メール")' if link else ""
                out = [
                    mail_id,
                    hit_id,
                    message_id,
                    excel_link,
                    date_str,
                    from_,
                    to_,
                    subj,
                    parttype,
                    matched,
                ]
                writer.writerow([self._sanitize_csv_field(v) for v in out])
        logging.info(f"[MailGrep] excel {csv_path}")

    def _write_xlsx(self, rows: list[list]) -> None:
        """
        XLSXで保存（openpyxl が必要）。
        - リンク列はExcelのHYPERLINK関数を使う（CSVと同じ）
        - 改行はCSVと同様に可視化（⏎）
        """

        wb: WorkbookType = Workbook()
        ws: Worksheet = cast(Worksheet, wb.active)
        ws.title = "results"

        headers: list[str] = [
            "mail_id",
            "hit_id",
            "message_id",
            "link",
            "Date",
            "From",
            "To",
            "Subject",
            "Matched Part",
            "Matched Line",
        ]
        ws.append(headers)
        # フィルタ機能を有効化
        ws.auto_filter.ref = ws.dimensions

        for row in rows:
            (
                _,
                mail_id,
                hit_id,
                message_id,
                link,
                date_str,
                from_,
                to_,
                subj,
                parttype,
                matched,
            ) = row

            mail_id: int
            hit_id: int
            message_id: str
            link: str
            date_str: str
            from_: str
            to_: str
            subj: str
            parttype: str
            matched: str

            excel_link: str = f'=HYPERLINK("{link}","メール")' if link else ""
            values: list[str | int] = [
                mail_id,
                hit_id,
                self._sanitize_csv_field(message_id),
                excel_link,
                self._sanitize_csv_field(date_str),
                self._sanitize_csv_field(from_),
                self._sanitize_csv_field(to_),
                self._sanitize_csv_field(subj),
                self._sanitize_csv_field(parttype),
                self._sanitize_csv_field(matched),
            ]
            ws.append(values)

        # 列幅自動調整
        try:
            for i, col in enumerate(
                ws.iter_cols(
                    min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
                ),
                1,
            ):
                max_length: int = 0
                for cell in col:
                    cell = cast(Cell, cell)
                    try:
                        cell_value: str = (
                            str(cell.value) if cell.value is not None else ""
                        )
                        max_length = max(max_length, len(cell_value))
                    except Exception:
                        pass
                col_letter: str = get_column_letter(i)
                ws.column_dimensions[col_letter].width = min(max_length + 2, 60)
        except Exception:
            logging.warning("[MailGrep] 列幅の自動調整に失敗しました。")
            pass

        # hit_id=1のみが表示されるようにデフォルトフィルタを設定（Excelで開いたとき）
        try:
            from openpyxl.worksheet.filters import (
                CustomFilter,
                CustomFilters,
                FilterColumn,
            )

            # hit_id列は2列目（index=1）
            filter_col = FilterColumn(
                colId=1,
                customFilters=CustomFilters(
                    customFilter=[CustomFilter(operator="equal", val="1")]
                ),
            )
            ws.auto_filter.filterColumn = [filter_col]
        except Exception:
            logging.warning("[MailGrep] デフォルトフィルタの設定に失敗しました。")
            pass
        xlsx_path = self._output_path.with_suffix(".xlsx")
        wb.save(xlsx_path)
        logging.info(f"[MailGrep] excel {xlsx_path}")


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
        default=None,
        help="出力CSVファイル名（未指定時は自動生成: results/<pattern先頭16>_タイムスタンプ.csv）",
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
    flags = re.IGNORECASE if args.ignore_case else 0

    finder = MailFolder(args.source)
    decoder = MailTextDecoder()
    pattern = SearchPattern(args.pattern, flags)
    output_path: Path = Path()
    if args.output:
        output_path = Path(args.output)
    else:
        unique_name = pattern.uid()
        out_dir = Path("results")
        out_dir.mkdir(parents=True, exist_ok=True)
        output_path = out_dir / f"{unique_name}.csv"

    exporter = MailCsvExporter(finder, decoder, pattern, output_path)
    exporter.process_all()


if __name__ == "__main__":
    MAIL_GREP_DEBUG = os.getenv("MAIL_GREP_DEBUG", "0") == "1"
    log_level = logging.DEBUG if MAIL_GREP_DEBUG else logging.INFO
    with SmartLogging(log_level) as env:
        env.set_stream_filter(True)
        try:
            main()
        except Exception as e:
            logging.error(f"[MailGrep] Unhandled exception: {e}")
            logging.error(
                "Unhandled exception at top-level:\n" + traceback.format_exc()
            )
            print("==== FATAL TRACEBACK ====")
            traceback.print_exc(file=sys.stdout)
            exit(1)
