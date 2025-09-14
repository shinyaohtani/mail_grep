from mail_profile import MailProfile
from mail_string_utils import CsvFieldText


class HitLine:
    def __init__(
        self,
        mail_keys: MailProfile,
        mail_id: int,
        hit_count: int,
        parttype: str,
        line: str,
    ):
        self.mail_keys = mail_keys
        self.mail_id = mail_id
        self.hit_count = hit_count
        self.parttype = parttype
        self.line = line.strip()

    def values(self) -> list[str | int]:
        ret = [
            self.mail_id,
            self.hit_count,
            self.mail_keys.message_id,
            self.mail_keys.excel_link,
            self.mail_keys.date_str,
            self.mail_keys.from_addr,
            self.mail_keys.to_addr,
            self.mail_keys.subj,
            self.parttype,
            self.line,
        ]
        return [CsvFieldText.sanitize(v) for v in ret]
