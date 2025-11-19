# mbox_classifier.py
from __future__ import annotations
from pathlib import Path
import plistlib, unicodedata, re
from typing import Iterable


EXCLUDE_TOKENS = {
    # drafts
    "draft",
    "drafts",
    "下書",
    "brouillon",
    "bozza",
    "entwurf",
    "borrador",
    "rascunho",
    "черновик",
    # trash / deleted / bin
    "trash",
    "deleted",
    "bin",
    "ゴミ箱",
    "削除済",
    "corbeille",
    "papierkorb",
    "papelera",
    "cestino",
    "корзина",
    # junk / spam
    "junk",
    "spam",
    "迷惑メール",
    "indesirable",
    "indésirable",
    "unerwünscht",
    "спам",
    # outbox
    "outbox",
    "送信トレイ",
    "posteausgang",
    # archive
    "archive",
    "アーカイブ",
    "archivo",
    "archivio",
    "archiv",
    # noise (任意除外)
    "rss",
    "メモ",
    "notes",
    "tasks",
    "タスク",
    "journal",
    "会話の履歴",
    "同期の問題",
    "recovered",
}
SENT_TOKENS = {
    "sent",
    "送信済み",
    "送信済みアイテム",
    "gesendet",
    "inviati",
    "enviados",
    "envoyes",
    "envoyés",
    "отправленные",
}
SPECIAL_ATTR_TOKENS = {
    "\\drafts",
    "\\junk",
    "\\trash",
    "\\deleted",
    "\\bin",
    "\\spam",
    "\\outbox",
    "\\archive",
}


class MboxClassifier:
    """mbox の役割を言語非依存で判定する分類器。"""

    # ---- 統合されたクラス内部ユーティリティ ----
    @classmethod
    def _norm(cls, s: str) -> str:
        s = unicodedata.normalize("NFKC", s).lower()
        return re.sub(r"[\s\W_]+", "", s, flags=re.UNICODE)

    @classmethod
    def _strings_from_plist(cls, plist_path: Path) -> list[str]:
        out: list[str] = []
        try:
            data = plistlib.loads(plist_path.read_bytes())
        except Exception:
            return out

        def walk(v):
            if isinstance(v, dict):
                for vv in v.values():
                    walk(vv)
            elif isinstance(v, (list, tuple, set)):
                for vv in v:
                    walk(vv)
            elif isinstance(v, bytes):
                try:
                    out.append(v.decode("utf-8", "ignore"))
                except Exception:
                    pass
            elif isinstance(v, str):
                out.append(v)

        walk(data)
        return out

    @classmethod
    def _hit(cls, tokens: set[str], *cands: Iterable[str]) -> bool:
        for seq in cands:
            for s in seq:
                ns = cls._norm(s)
                if any(tok in ns for tok in tokens):
                    return True
        return False

    # ---- 公開メソッド ----
    def is_excluded(self, mbox_dir: Path) -> bool:
        info = mbox_dir / "Info.plist"
        if info.exists():
            strs = self._strings_from_plist(info)
            # ① 特別属性判定（最優先）
            if self._hit(SPECIAL_ATTR_TOKENS, strs):
                return True
            # ② 下書き/迷惑/削除/ノイズ判定
            if self._hit(EXCLUDE_TOKENS, strs):
                return True

        # ③ フォールバック: 名前ベース
        names = [mbox_dir.name.replace(".mbox", ""), mbox_dir.parent.name]
        hit = self._hit(EXCLUDE_TOKENS, names)
        return hit

    def is_sent(self, mbox_dir: Path) -> bool:
        info = mbox_dir / "Info.plist"
        if info.exists():
            strs = self._strings_from_plist(info)
            if self._hit({"\\sent"}, strs):
                return True
            if self._hit(SENT_TOKENS, strs):
                return True

        names = [mbox_dir.name.replace(".mbox", ""), mbox_dir.parent.name]
        return self._hit(SENT_TOKENS, names)
