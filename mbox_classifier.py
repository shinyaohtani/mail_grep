# mbox_classifier.py
from __future__ import annotations
from pathlib import Path
import plistlib, unicodedata, re
from typing import Iterable

# 言語非依存トークン（正規化後の連結小文字文字列に部分一致）
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


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", s).lower()
    return re.sub(r"[\s\W_]+", "", s, flags=re.UNICODE)


def _strings_from_plist(plist_path: Path) -> list[str]:
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


def _hit(tokens: set[str], *cands: Iterable[str]) -> bool:
    for seq in cands:
        for s in seq:
            ns = _norm(s)
            if any(tok in ns for tok in tokens):
                return True
    return False


class MboxClassifier:
    """mbox の“役割”を言語非依存で判定する小さな分類器。"""

    def is_excluded(self, mbox_dir: Path) -> bool:
        """Drafts/Trash/Junk/Outbox/Archive/RSS/メモ…等なら True。"""
        info = mbox_dir / "Info.plist"
        if info.exists():
            strs = _strings_from_plist(info)
            if _hit(SPECIAL_ATTR_TOKENS, strs):  # \Drafts 等を最優先
                return True
            if _hit(EXCLUDE_TOKENS, strs):
                return True
        # フォールバック：ディレクトリ名・親名でも判定
        names = [mbox_dir.name.replace(".mbox", ""), mbox_dir.parent.name]
        return _hit(EXCLUDE_TOKENS, names)

    def is_sent(self, mbox_dir: Path) -> bool:
        """“送信済み”と判断できれば True"""
        info = mbox_dir / "Info.plist"
        if info.exists():
            if _hit({"\\sent"}, _strings_from_plist(info)):
                return True
            if _hit(SENT_TOKENS, _strings_from_plist(info)):
                return True
        names = [mbox_dir.name.replace(".mbox", ""), mbox_dir.parent.name]
        return _hit(SENT_TOKENS, names)
