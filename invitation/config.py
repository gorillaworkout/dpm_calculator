"""Settings for DPM Invitation, from DW_* environment variables at start-up.

DW_INVITATION_DB    SQLite file (default: invitation.db next to users.db)
DW_INVITATION_DIR   uploaded invitation images (default: invitation_files/ next to users.db)
DW_PUBLIC_BASE_URL  base of the public link (default https://dw-calculator.gorillaworkout.id),
                    shared with Partner Approval

Both paths live next to users.db, like Partner Approval's approval.db and
approval_files/, so a deploy that copies the code files never touches them.
"""
import os
from dataclasses import dataclass
from pathlib import Path

HERE = Path(__file__).resolve().parent


@dataclass
class Settings:
    db: Path
    files: Path
    base_url: str


def load():
    env = os.environ.get
    auth_db = Path(env("DW_AUTH_DB") or HERE.parent / "users.db")
    return Settings(
        db=Path(env("DW_INVITATION_DB") or auth_db.with_name("invitation.db")).expanduser(),
        files=Path(env("DW_INVITATION_DIR") or auth_db.with_name("invitation_files")).expanduser(),
        base_url=(env("DW_PUBLIC_BASE_URL") or "https://dw-calculator.gorillaworkout.id").rstrip("/"),
    )
