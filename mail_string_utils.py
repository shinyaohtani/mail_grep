from __future__ import annotations
from email.header import decode_header
from typing import Any, Iterable
import logging


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
        cleaned: str = EncodedHeader.remove_crlf(value)
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

    @staticmethod
    def remove_crlf(value: str) -> str:
        if not value:
            return ""
        # 既存仕様：CR除去、LFはスペースに
        return value.replace("\r", "").replace("\n", " ")
