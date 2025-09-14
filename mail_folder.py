from pathlib import Path

class MailFolder:
    def __init__(self, root_dir: Path):
        self._root_dir = root_dir

    def mail_paths(self) -> list[Path]:
        return list(self._root_dir.rglob("*.emlx"))
