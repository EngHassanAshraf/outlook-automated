import logging
import re
from typing import Any, Optional, TYPE_CHECKING

# to correctly display arabic language in terminal
from bidi.algorithm import get_display

from datetime import date
from pathlib import Path

from config_manager import config

if TYPE_CHECKING:
    from progress_ui import ProgressUI

# Module-level UI reference. Set by main.py before processing begins so that
# Mail.move_mail() can push events into the live panel instead of printing.
_ui: Optional["ProgressUI"] = None


def set_ui(ui: Optional["ProgressUI"]) -> None:
    """Register the active ProgressUI instance for this module."""
    global _ui
    _ui = ui


class AttachmentPath:
    """Handles building attachment file paths."""

    def attachment_path(
        self,
        output_dir: Path,
        category: str,
        compound: str,
        month: str,
        sub_category: Optional[str] = None,
    ) -> str:
        """
        Build the full folder path for saving an attachment.

        Args:
            output_dir:   Base output directory.
            category:     Email category.
            compound:     Sender compound name.
            month:        Month string, e.g. "4. April".
            sub_category: Optional sub-category.

        Returns:
            Full folder path string.
        """
        week_number: Optional[int] = None

        if "Tech" in category or category == "التشغيلات الاسبوعية":
            week_number = (int(date.today().day) - 1) // 7 + 1

        sub_category_str = sub_category if sub_category else ""
        week_str = f"week {week_number}" if week_number else ""

        path = f"{output_dir}\\{category}\\{compound}\\{sub_category_str}\\{month}\\{week_str}"
        return path.replace("\\\\", "\\")


class Mail:
    """Handles email message operations."""

    def __init__(self, message: Any):
        self.mail = message

    def get_mail_attachments(self) -> Any:
        """Return the attachments collection of this email."""
        return self.mail.attachments

    def move_mail(self, folder: Any) -> bool:
        """
        Move this email to *folder*.

        Pushes a ☑ line into the live UI panel (or falls back to logging
        if the UI is not active).

        Returns:
            True on success, False on failure.
        """
        try:
            subject = str(self.mail.subject)
            self.mail.Move(folder)
            logging.info(f"Archived: {subject}")
            if _ui is not None:
                _ui.notify(f"☑  {get_display(subject)}")
            return True
        except Exception as exc:
            logging.error(f"Error moving email: {exc}")
            if _ui is not None:
                _ui.error(f"✗  Move failed: {exc}")
            return False

    def mark_read(self) -> bool:
        """Mark this email as read."""
        try:
            self.mail.Unread = False
            self.mail.Save()
            logging.info(f"Marked as read: {self.mail.subject}")
            return True
        except Exception as exc:
            logging.error(f"Error marking email as read: {exc}")
            return False

    def is_read(self) -> bool:
        """Return True if the email has been read."""
        return not self.mail.unread


class Attachment:
    """Handles email attachment operations."""

    _DATE_PATTERNS = [
        re.compile(r'(?P<year>\d{4})[._-](?P<month>\d{1,2})[._-](?P<day>\d{1,2})'),  # YYYY-MM-DD
        re.compile(r'(?P<day>\d{1,2})[._-](?P<month>\d{1,2})[._-](?P<year>\d{4})'),  # DD-MM-YYYY
        re.compile(r'(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})'),                 # YYYYMMDD
        re.compile(r'(?P<day>\d{2})(?P<month>\d{2})(?P<year>\d{4})'),                 # DDMMYYYY
    ]

    def __init__(self, attachment: Any):
        self.attachment = attachment

    def is_ignored(self) -> bool:
        """Return True if this attachment is on the ignore list."""
        return str(self.attachment.filename) in config.get_ignored_files()

    def accepted_type(self) -> bool:
        """Return True if the attachment's file extension is accepted."""
        filename = str(self.attachment.filename)
        if "." not in filename:
            return False
        return filename.rsplit(".", 1)[-1].lower() in config.get_accepted_types()

    def attachment_month(self, item: Any) -> str:
        """
        Determine the month for this attachment.

        Tries common date patterns in the filename; falls back to ReceivedTime.

        Returns:
            Month string in the format "N. MonthName" (e.g. "4. April").
        """
        filename = str(self.attachment.filename)
        month_number: Optional[int] = None

        for pattern in self._DATE_PATTERNS:
            match = pattern.search(filename)
            if match:
                try:
                    candidate = int(match.group('month'))
                    if 1 <= candidate <= 12:
                        month_number = candidate
                        break
                except (ValueError, IndexError):
                    continue

        if month_number is None:
            month_number = item.ReceivedTime.month

        month_name = date(2000, month_number, 1).strftime("%B")
        return f"{month_number}. {month_name}"

    def attachment_folder(self, path: str) -> str:
        """
        Create the folder and return a collision-safe file path.

        Appends (1), (2), … before the extension if the file already exists.
        """
        path_obj = Path(path)
        path_obj.mkdir(parents=True, exist_ok=True)

        filename = str(self.attachment.filename)
        target = path_obj / filename

        if target.exists():
            stem, suffix = target.stem, target.suffix
            counter = 1
            while target.exists():
                target = path_obj / f"{stem}({counter}){suffix}"
                counter += 1
            msg = f"⚠  Collision: saved as {target.name}"
            logging.warning(msg)
            if _ui is not None:
                _ui.warn(msg)

        return str(target)

    def save_attachment(self, file_path: str) -> bool:
        """Save this attachment to *file_path*."""
        try:
            self.attachment.SaveAsFile(file_path)
            logging.info(f"Saved: {Path(file_path).name}")
            return True
        except Exception as exc:
            logging.error(f"Error saving '{file_path}': {exc}")
            if _ui is not None:
                _ui.error(f"✗  Save failed: {Path(file_path).name}")
            return False
