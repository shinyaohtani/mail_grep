from datetime import datetime
from dataclasses import dataclass


@dataclass(frozen=True)
class MailProfile:
    message_id: str
    date_str: str
    date_dt: datetime | None
    link: str
    subj: str
    from_addr: str
    to_addr: str

    @property
    def excel_link(self) -> str:
        return f'=HYPERLINK("{self.link}","メール")' if self.link else ""
