from __future__ import annotations

from bs4 import BeautifulSoup
from datetime import datetime
from email import policy
from email.message import Message
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Generator
from urllib.parse import quote
import email
import logging

from mail_profile import MailProfile
from mail_string_utils import AnyText, EncodedHeader
from search_pattern import SearchPattern


class _MailBlob:
    """'.emlx' のバイト実体を保持するモノ。"""

    def __init__(self, mail_abspath: Path):
        self._mail_abspath: Path = mail_abspath

    def bytes(self) -> bytes:
        raw: bytes = self._mail_abspath.read_bytes()
        return raw[raw.find(b"\n") + 1 :] if raw[:1].isdigit() else raw

    def message(self) -> Message:
        data: bytes = self.bytes()
        if data[:1].isdigit():
            nl: int = data.find(b"\n")
            if nl != -1:
                data = data[nl + 1 :]
        parser: BytesParser = BytesParser(policy=policy.default)
        try:
            return parser.parsebytes(data)
        except Exception:
            return email.message_from_bytes(data, policy=policy.default)


class _MailHeaders:
    """メールのヘッダという“モノ”。"""

    def __init__(self, msg: Message):
        self._msg: Message = msg

    def id_str(self) -> str:
        try:
            v: str | None = AnyText.to_str(self._msg.get("Message-ID"))
            return EncodedHeader.decode(v) if v else ""
        except Exception:
            return ""

    def date_str(self) -> str:
        try:
            v: str | None = AnyText.to_str(self._msg.get("Date"))
            if not v:
                return ""
            dt: datetime = parsedate_to_datetime(v)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            try:
                v2: str | None = AnyText.to_str(self._msg.get("Date"))
                return EncodedHeader.decode(v2 or "")
            except Exception:
                return ""

    def date_dt(self) -> datetime | None:
        try:
            v: str | None = AnyText.to_str(self._msg.get("Date"))
            return parsedate_to_datetime(v) if v else None
        except Exception:
            return None

    def link(self) -> str:
        mid: str = self.id_str()
        if not mid:
            return ""
        s: str = mid.strip()
        if not s.startswith("<"):
            s = f"<{s}>"
        return f"message:{quote(s, safe='')}"

    def lines(self) -> list[str]:
        result: list[str] = []
        for k in ("Subject", "From", "To", "Date"):
            try:
                v: str | None = AnyText.to_str(self._msg.get(k))
                if v is None:
                    continue
                decoded: str = EncodedHeader.decode(v)
                decoded = decoded.replace("\r", "").replace("\n", " ")
                result.append(f"{k}: {decoded}")
            except Exception as e:
                logging.warning(f"[MailGrep] Could not parse {k}: {e}")
                result.append(f"{k}: [INVALID HEADER]")
        return result

    def subj(self) -> str:
        v: str | None = AnyText.to_str(self._msg.get("Subject"))
        return EncodedHeader.decode(v) if v else ""

    def from_addr(self) -> str:
        v: str | None = AnyText.to_str(self._msg.get("From"))
        return EncodedHeader.decode(v) if v else ""

    def to_addr(self) -> str:
        v: str | None = AnyText.to_str(self._msg.get("To"))
        return EncodedHeader.decode(v) if v else ""


class _MailBody:
    """メールの本文という“モノ”。"""

    def __init__(self, msg: Message):
        self._msg: Message = msg

    def lines(self) -> list[tuple[str, str]]:
        result: list[tuple[str, str]] = []
        for part in self._iter_text_parts():
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            text: str = self._decode_payload(part, payload)
            ctype: str = part.get_content_type()
            if ctype == "text/html":
                result.extend(self._html_variants(text))
            else:
                for line in text.splitlines():
                    if line.strip():
                        result.append((line, ctype))
        return result

    def _iter_text_parts(self) -> Generator[Message, None, None]:
        if hasattr(self._msg, "is_multipart") and self._msg.is_multipart():
            for p in self._msg.walk():
                if p.get_content_type().startswith("text/"):
                    yield p
        else:
            yield self._msg

    def _decode_payload(self, part: Message, payload) -> str:
        charset: str = part.get_content_charset() or "utf-8"
        try:
            return bytes(payload).decode(charset, errors="replace")
        except Exception:
            return bytes(payload).decode("utf-8", errors="replace")

    def _html_variants(self, html: str) -> list[tuple[str, str]]:
        out: list[tuple[str, str]] = []
        out.append((html, "text/html"))
        soup: BeautifulSoup = BeautifulSoup(html, "html.parser")
        for tag in soup(["head", "script", "style", "meta", "title", "link"]):
            tag.decompose()
        text_only: str = soup.get_text(separator="\n", strip=True)
        for line in text_only.splitlines():
            if line.strip():
                out.append((line, "text/html_textonly"))
        concat: str = "".join(soup.stripped_strings)
        if concat:
            out.append((concat, "text/html_concat"))
        return out


class MailMessage:
    """既存 API を保つ外側の“器”。"""

    def __init__(self, mail_abspath: Path):
        self._mail_abspath: Path = mail_abspath
        blob: _MailBlob = _MailBlob(mail_abspath)
        self._msg: Message = blob.message()
        self._headers: _MailHeaders = _MailHeaders(self._msg)
        self._body: _MailBody = _MailBody(self._msg)
        self._profile: MailProfile = self._create_profile()

    def key_profile(self) -> MailProfile:
        return self._profile

    def extract(self, pattern: SearchPattern) -> list[tuple[str, str]]:
        matches: list[tuple[str, str]] = []
        for line in self._header_lines():
            if pattern.check_line(line):
                matches.append(("header", line))
        for line, parttype in self._body_lines():
            if parttype not in ("text/plain", "text/html_textonly"):
                continue
            if pattern.check_line(line):
                matches.append((parttype, line))
        return matches

    def _header_lines(self) -> list[str]:
        return self._headers.lines()

    def _body_lines(self) -> list[tuple[str, str]]:
        return self._body.lines()

    def _create_profile(self) -> MailProfile:
        return MailProfile(
            message_id=self._headers.id_str(),
            date_str=self._headers.date_str(),
            date_dt=self._headers.date_dt(),
            link=self._headers.link(),
            subj=self._headers.subj(),
            from_addr=self._headers.from_addr(),
            to_addr=self._headers.to_addr(),
        )
