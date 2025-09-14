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
from mail_string_utils import MailStringUtils
from search_pattern import SearchPattern


class MailMessage:
    def __init__(self, mail_path: Path):
        self._mail_path = mail_path
        self._msg: Message = self._load()
        self._profile = self.profile()

    def key_profile(self) -> MailProfile:
        return self._profile
    
    def extract(self, pattern: SearchPattern) -> list[tuple[str, str]]:
        matches: list[tuple[str, str]] = []
        # 1) ヘッダー行はこれまでどおり検索
        for line in self.header_lines():
            if pattern.check_line(line):
                matches.append(("header", line))

        # 2) 本文行は text/plain と text/html_textonly のみ対象
        for line, parttype in self.body_lines():
            if parttype not in ("text/plain", "text/html_textonly"):
                continue
            if pattern.check_line(line):
                matches.append((parttype, line))

        return matches

    @property
    def _id_str(self) -> str:
        try:
            v = self._msg.get("Message-ID")
            v = MailStringUtils.stringify(v)
            return MailStringUtils.decode_header(v) if v else ""
        except Exception:
            return ""

    @property
    def _date_str(self) -> str:
        try:
            v = self._msg.get("Date")
            v = MailStringUtils.stringify(v)
            if not v:
                return ""
            dt = parsedate_to_datetime(v)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            try:
                v = self._msg.get("Date")
                v = MailStringUtils.stringify(v)
                return MailStringUtils.decode_header(v or "")
            except Exception:
                return ""

    @property
    def _date_dt(self) -> datetime | None:
        """並べ替え用に Date を datetime で返す（失敗時は None）"""
        try:
            v = self._msg.get("Date")
            v = MailStringUtils.stringify(v)
            if not v:
                return None
            return parsedate_to_datetime(v)
        except Exception:
            return None

    @property
    def _link_str(self) -> str:
        """
        Mail.app で開けるリンクを生成。
        形式: message:%3Cmessage-id%3E  （< と > は URL エンコード）
        """
        message_id = self._id_str
        if not message_id:
            return ""
        mid = message_id.strip()
        if not mid.startswith("<"):
            mid = f"<{mid}>"
        # 角括弧や@を含めて完全にエンコード（Excel でもクリック可能）
        return f"message:{quote(mid, safe='')}"

    @property
    def _subject_str(self) -> str:
        subj = MailStringUtils.decode_header(
            MailStringUtils.stringify(self._msg.get("Subject"))
        )
        return subj

    @property
    def _from_addr(self) -> str:
        from_ = MailStringUtils.decode_header(
            MailStringUtils.stringify(self._msg.get("From"))
        )
        return from_

    @property
    def _to_addr(self) -> str:
        to_ = MailStringUtils.decode_header(
            MailStringUtils.stringify(self._msg.get("To"))
        )
        return to_

    def profile(self) -> MailProfile:
        return MailProfile(
            message_id=self._id_str,
            date_str=self._date_str,
            date_dt=self._date_dt,
            link=self._link_str,
            subj=self._subject_str,
            from_addr=self._from_addr,
            to_addr=self._to_addr,
        )

    def header_lines(self) -> list[str]:
        result = []
        for k in ("Subject", "From", "To", "Date"):
            try:
                v = self._msg.get(k)
                v = MailStringUtils.stringify(v)
                if v is None:
                    continue
                decoded = MailStringUtils.decode_header(v)
                decoded = decoded.replace("\r", "").replace("\n", " ")
                result.append(f"{k}: {decoded}")
            except Exception as e:
                logging.warning(f"[MailGrep] Could not parse {k}: {e}")
                result.append(f"{k}: [INVALID HEADER]")
        return result

    def body_lines(self) -> list[tuple[str, str]]:
        result = []
        for part in self._iter_text_parts():
            payload = part.get_payload(decode=True)
            if not payload:
                continue

            charset: str = part.get_content_charset() or "utf-8"
            html: str
            try:
                html = bytes(payload).decode(charset, errors="replace")
            except Exception:
                html = bytes(payload).decode("utf-8", errors="replace")

            ctype: str = part.get_content_type()
            if ctype == "text/html":
                # ① 生HTMLをそのまま
                result.append((html, "text/html"))

                # ② BeautifulSoupでパース
                soup = BeautifulSoup(html, "html.parser")
                #    不要タグをまるごと削除
                for tag in soup(["head", "script", "style", "meta", "title", "link"]):
                    tag.decompose()

                # ③ テキストのみ改行区切りで取得
                text_only: str = soup.get_text(separator="\n", strip=True)
                for line in text_only.splitlines():
                    if line.strip():
                        result.append((line, "text/html_textonly"))

                # ④ さらに全文連結版も追加
                concat: str = "".join(soup.stripped_strings)
                if concat:
                    result.append((concat, "text/html_concat"))

            else:
                # text/plain
                text: str = html
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

    @staticmethod
    def _read_emlx(mail_path: Path) -> bytes:
        raw = mail_path.read_bytes()
        return raw[raw.find(b"\n") + 1 :] if raw[0:1].isdigit() else raw

    def _load(self) -> Message:
        """
        .emlx の長さ行をスキップしたあと、
        ヘッダー＋本文をそのまま BytesParser に渡す
        """
        raw_bytes = self._read_emlx(self._mail_path)
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
