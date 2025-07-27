import csv
import email
from email import policy
from email.header import decode_header
from pathlib import Path


class EmlxPathCollector:
    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def collect(self) -> list[Path]:
        return list(self._root_dir.rglob("*.emlx"))


class MailHeaderExtractor:
    def extract(self, emlx_path: Path) -> tuple[str, str, str, str] | None:
        try:
            raw = self._read_raw_emlx(emlx_path)
            msg = email.message_from_bytes(raw, policy=policy.default)
            return (
                self._decode_header(msg["Subject"]),
                self._decode_header(msg["From"]),
                self._decode_header(msg["To"]),
                self._decode_header(msg["Date"]),
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


class CsvWriter:
    def __init__(self, output_path: Path):
        self._output_path = output_path

    def write(self, rows: list[tuple[str, str, str, str]]) -> None:
        with open(self._output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Subject", "From", "To", "Date"])
            writer.writerows(rows)


class MailSummaryExporter:
    def __init__(self, source_dir: Path, output_path: Path):
        self._collector = EmlxPathCollector(source_dir)
        self._extractor = MailHeaderExtractor()
        self._writer = CsvWriter(output_path)

    def export(self) -> None:
        rows = []
        for path in self._collector.collect():
            result = self._extractor.extract(path)
            if result:
                rows.append(result)
        self._writer.write(rows)


if __name__ == "__main__":
    source = Path.home() / "Library" / "Mail" / "V10"  # バージョン確認要
    output = Path("output_mail_summary.csv")
    MailSummaryExporter(source, output).export()
