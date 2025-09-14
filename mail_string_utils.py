from __future__ import annotations

from email.header import decode_header
import logging
from typing import Any, Iterable


class CsvFieldText:
    """CSVの1セルに収まる文字列表現。"""

    @staticmethod
    def sanitize(value: Any) -> str:
        s: str = CsvFieldText._to_str(value)
        # 既存仕様：CR削除、LFは ⏎ に置換
        return s.replace("\r", "").replace("\n", "⏎")

    @staticmethod
    def _to_str(value: Any) -> str:
        if value is None:
            return ""
        return str(value)


class AnyText:
    """任意オブジェクトの“文字列表現”を担うテキスト塊。"""

    @staticmethod
    def to_str(v: str | bytes | None) -> str:
        if v is None:
            return ""
        # email.headerregistry.Address / AddressHeader 等の .addresses を許容
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


class EncodedHeader:
    """エンコード断片から成るメールヘッダ“テキスト”。"""

    @staticmethod
    def decode(value: str | None) -> str:
        if value is None:
            return ""
        cleaned: str = RawHeaderText.remove_crlf(value)
        try:
            parts: Iterable[tuple[bytes | str, str | None]] = decode_header(cleaned)
        except Exception:
            return ""

        out: list[str] = []
        for text, enc in parts:
            if isinstance(text, bytes):
                status: str = "noErr"
                if enc:
                    try:
                        decoded: str = text.decode(enc, errors="strict")
                        out.append(decoded)
                        continue
                    except Exception as e:
                        status = f"decode_header: enc='{enc}' decode failed ({e}), trying utf-8 ..."
                out.append(EncodedHeader._fallback(text, status))
            else:
                out.append(str(text))
        return "".join(out)

    @staticmethod
    def _fallback(text: bytes, status: str) -> str:
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


class RawHeaderText:
    """CRLFの扱いやヘッダ部抽出を担う“生ヘッダ文字列”。"""

    @staticmethod
    def remove_crlf(value: str) -> str:
        if not value:
            return ""
        # 既存仕様：CR除去、LFはスペースに
        return value.replace("\r", "").replace("\n", " ")

    @staticmethod
    def sanitize_section(header_str: str) -> str:
        sanitized: list[str] = []
        prev_colon: bool = False
        for line in header_str.splitlines():
            if ":" in line:
                key: str
                val: str
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

    @staticmethod
    def decode_bytes(raw: bytes) -> str:
        return raw.decode("utf-8", errors="ignore")

    @staticmethod
    def headers_section(raw_str: str) -> str:
        lines: list[str] = []
        for line in raw_str.splitlines():
            if line.strip() == "":
                break
            lines.append(line)
        return "\n".join(lines)


class MailStringUtils:
    """
    互換APIファサード。
    既存コードと同じメソッド名・同じシグネチャ・同じ振る舞いを維持する。
    """

    @staticmethod
    def sanitize_csv_field(value: Any) -> str:
        return CsvFieldText.sanitize(value)

    @staticmethod
    def stringify(v: str | bytes | None) -> str:
        return AnyText.to_str(v)

    @staticmethod
    def decode_header(value: str | None) -> str:
        return EncodedHeader.decode(value)

    @staticmethod
    def remove_crlf(value: str) -> str:
        return RawHeaderText.remove_crlf(value)

    # 以降は“旧private相当の内部機能”は RawHeaderText / EncodedHeader に集約。
    # もし既存コードがこれらの private 名を直接参照していない限り互換性は維持されます。
