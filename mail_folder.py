from pathlib import Path
from mbox_classifier import MboxClassifier


class MailFolder:
    def __init__(self, root_dir: Path, only_sent: bool = False):
        self._root_dir = root_dir
        self._only_sent = only_sent
        self._clf = MboxClassifier()

    def mail_paths(self) -> list[Path]:
        emlxs: list[Path] = []
        for mbox in self._root_dir.rglob("*.mbox"):
            if self._only_sent:
                if not self._clf.is_sent(mbox):
                    continue
            else:
                if self._clf.is_excluded(mbox):
                    print(f"Exclude: {mbox}")
                    continue
            emlxs.extend(mbox.rglob("*.emlx"))
        return sorted(emlxs, key=lambda p: p.stat().st_mtime, reverse=True)
