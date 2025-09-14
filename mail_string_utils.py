from email.header import decode_header
import logging

from dataclasses import dataclass


class MailStringUtils:
    @staticmethod
    def sanitize_csv_field(value) -> str:
        if value is None:
            return ""
        s = str(value)
        return s.replace("\r", "").replace("\n", "⏎")

    @staticmethod
    def stringify(v: str | bytes | None) -> str:
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

    @staticmethod
    def decode_header(value: str | None) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            value = MailStringUtils.remove_crlf(value)
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
                out.append(MailStringUtils._decode_header_fallback(text, status))
            else:
                out.append(str(text))
        return "".join(out)

    @staticmethod
    def _decode_header_fallback(text: bytes, status: str) -> str:
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

    @staticmethod
    def _sanitize_header_section(header_str: str) -> str:
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

    @staticmethod
    def _decode_bytes(raw: bytes) -> str:
        return raw.decode("utf-8", errors="ignore")

    @staticmethod
    def _headers_section(raw_str: str) -> str:
        lines = []
        for line in raw_str.splitlines():
            if line.strip() == "":
                break
            lines.append(line)
        return "\n".join(lines)
