from pathlib import Path
import argparse
import logging
import os
import re
import sys
import traceback

from hit_line import HitLine
from hit_report import HitReport
from mail_folder import MailFolder
from mail_message import MailMessage
from mail_profile import MailProfile
from search_pattern import SearchPattern
from smart_logging import SmartLogging


class MailGrepApp:
    def __init__(
        self,
        mail_storage: MailFolder,
        pattern: SearchPattern,
        output_path: Path | None,
    ):
        self._storage = mail_storage
        self._pattern = pattern
        if output_path is None:
            unique_name = pattern.unique_name
            out_dir = Path("results")
            out_dir.mkdir(parents=True, exist_ok=True)
            output_path = out_dir / f"{unique_name}.csv"
        self._output_path = output_path

    def run(self):
        report: HitReport = HitReport()
        try:
            for mail_id, path in enumerate(self._storage.mail_paths(), 1):
                try:
                    message: MailMessage = MailMessage(path)
                    mail_keys: MailProfile = message.key_profile()
                    for hit_count, (parttype, line) in enumerate(
                        message.extract(self._pattern), 1
                    ):
                        if hit_count == 1:
                            print(
                                f"✓ [{parttype}] hit! ({mail_keys.date_str}) {self.line_preview(line)}"
                            )
                        else:
                            print(f"✓ [{parttype}] {self.line_preview(line)}")
                        hit = HitLine(
                            mail_keys,  # use date_dt as a sort key (datetime or None)
                            mail_id,  # mail_id
                            hit_count,  # hit_id
                            parttype,  # Matched Part
                            line.strip(),  # Matched Line
                        )
                        report.append_hit_line(hit)
                except Exception as e:
                    logging.warning(f"[MailGrep] Skipped {path}: {e}")
        except KeyboardInterrupt:
            print()
            logging.warning("[MailGrep]中断されました。ここまでの結果を保存します。")
        finally:
            # 1) Date 降順でソート（None は末尾）
            report.sort()

            # 拡張子が何であろうと2種類とも保存する
            report.store(self._output_path.with_suffix(".xlsx"))
            report.store(self._output_path.with_suffix(".csv"))

            # 3) メール件数をログに表示（hit_id == 1 の行のみカウント）
            logging.info(f"[MailGrep] {report.mail_count()}件を保存しました。")

    @staticmethod
    def line_preview(line: str) -> str:
        max_len = 50
        preview = line.strip()
        if len(preview) > max_len:
            preview = preview[:max_len] + "..."
        return preview


class AppArguments:
    def __init__(self):
        self.parser: argparse.ArgumentParser = self._create_parser()
        self.args: argparse.Namespace | None = None

    def parse(self) -> argparse.Namespace:
        self.args = self.parser.parse_args()
        if self.args is None:
            raise ValueError("Failed to parse arguments")
        return self.args

    @staticmethod
    def _create_parser() -> argparse.ArgumentParser:
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


def main():
    app_args = AppArguments()
    args = app_args.parse()

    storage = MailFolder(args.source)

    flags = re.IGNORECASE if args.ignore_case else 0
    pattern = SearchPattern(args.pattern, flags)

    app = MailGrepApp(storage, pattern, args.output)
    app.run()


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
