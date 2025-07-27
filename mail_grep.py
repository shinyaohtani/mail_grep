import csv
import email
import re
import sys
from email import policy
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path


class EmlxPathCollector:
    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def collect(self) -> list[Path]:
        return list(self._root_dir.rglob("*.emlx"))


class MailHeaderExtractor:
    def extract(self, emlx_path: Path) -> tuple[str, str, str, str, str] | None:
        try:
            raw = self._read_raw_emlx(emlx_path)
            raw_str = raw.decode("utf-8", errors="ignore")
            match_line = self._search_target(raw_str)
            if not match_line:
                return None

            header_str = self._extract_headers(raw_str)
            msg = email.message_from_string(header_str, policy=policy.default)

            return (
                self._decode_header(msg["Subject"]),
                self._decode_header(msg["From"]),
                self._decode_header(msg["To"]),
                self._decode_date(msg["Date"]),
                match_line.strip(),
            )
        except Exception:
            return None

    def _read_raw_emlx(self, path: Path) -> bytes:
        raw = path.read_bytes()
        if raw[0:1].isdigit():
            return raw[raw.find(b"\n") + 1 :]
        return raw

    def _decode_header(self, value: str | None) -> str:
        if value is None:
            return ""
        parts = decode_header(value)
        return "".join(
            str(p[0], p[1] or "utf-8") if isinstance(p[0], bytes) else str(p[0])
            for p in parts
        )

    def _decode_date(self, value: str | None) -> str:
        if value is None:
            return ""
        try:
            dt = parsedate_to_datetime(value)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return value

    def _extract_headers(self, raw_str: str) -> str:
        lines = raw_str.splitlines()
        header_lines = []
        for line in lines:
            if line.strip() == "":
                break
            header_lines.append(line)
        return "\n".join(header_lines)

    def _search_target(self, raw_str: str) -> str | None:
        for line in raw_str.splitlines():
            if self._pattern.search(line):
                return line
        return None

    def set_pattern(self, pattern: str):
        self._pattern = re.compile(pattern)


class CsvWriter:
    def __init__(self, output_path: Path):
        self._output_path = output_path

    def write(self, rows: list[tuple[str, str, str, str, str]]) -> None:
        with open(self._output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Subject", "From", "To", "Date", "Matched Line"])
            writer.writerows(rows)


class MailGrepExporter:
    def __init__(self, source_dir: Path, output_path: Path, pattern: str):
        self._collector = EmlxPathCollector(source_dir)
        self._extractor = MailHeaderExtractor()
        self._extractor.set_pattern(pattern)
        self._writer = CsvWriter(output_path)

    def export(self) -> None:
        rows = []
        for path in self._collector.collect():
            result = self._extractor.extract(path)
            if result:
                print("✓", result[0], "←", result[4])
                rows.append(result)
        self._writer.write(rows)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python -m mail_grep '<pattern>'")
        sys.exit(1)

    pattern = sys.argv[1]
    source = Path.home() / "Library" / "Mail" / "V10"  # 適宜修正
    output = Path("output_mail_summary.csv")

    MailGrepExporter(source, output, pattern).export()
