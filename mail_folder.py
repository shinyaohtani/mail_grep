from pathlib import Path


class MailFolder:
    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def mail_paths(self) -> list[Path]:
        files = list(self._root_dir.rglob("*.emlx"))
        return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)
