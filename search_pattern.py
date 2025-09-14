from datetime import datetime
import re
import unicodedata

from mail_message import MailMessage


class SearchPattern:
    def __init__(self, egrep_pattern: str, flags=0):
        self._egrep_pattern = egrep_pattern
        python_pattern = self._egrep_to_python_regex(self._egrep_pattern)
        self._pattern = re.compile(python_pattern, flags | re.DOTALL)

    # --- egrep to Python正規表現（必要なら改良して下さい） ---
    @staticmethod
    def _egrep_to_python_regex(pattern: str) -> str:
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

    def match_mail(self, message: MailMessage) -> list[tuple[str, str]]:
        matches: list[tuple[str, str]] = []
        # 1) ヘッダー行はこれまでどおり検索
        for line in message.header_lines():
            if self._pattern.search(line):
                matches.append(("header", line))

        # 2) 本文行は text/plain と text/html_textonly のみ対象
        for line, parttype in message.body_lines():
            if parttype not in ("text/plain", "text/html_textonly"):
                continue
            if self._pattern.search(line):
                matches.append((parttype, line))

        return matches

    def unique_name(self) -> str:
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
