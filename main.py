import logging
import sys
from pathlib import Path
from datetime import date
from typing import Tuple, Optional, List, Any

from application import Connection, Folder
from message import Mail, Attachment, AttachmentPath, set_ui
from config_manager import config
from progress_ui import ProgressUI, SuppressConsoleHandler

import psutil

# Configure logging
def setup_logging(ui: Optional[ProgressUI] = None, file_only: bool = False) -> None:
    """
    Configure logging handlers.

    Args:
        ui:        When provided, attaches a SuppressConsoleHandler (no-op).
        file_only: When True, skips console entirely — only the file handler
                   is attached. Used from main() so nothing leaks to the
                   terminal before or during the live UI.
    """
    logging_config = config.get_logging_config()
    level = getattr(logging, logging_config.get('level', 'INFO').upper())
    log_format = logging_config.get('format', '%(asctime)s - %(levelname)s - %(message)s')

    handlers: list = []

    if not file_only and ui is None and logging_config.get('console', True):
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(logging.Formatter(log_format))
        handlers.append(console_handler)

    if logging_config.get('file', 'outlook_auto.log'):
        file_handler = logging.FileHandler(logging_config['file'], encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(logging.Formatter(log_format))
        handlers.append(file_handler)

    logging.basicConfig(level=level, format=log_format, handlers=handlers, force=True)


# Logging is configured inside main() before any work begins.
# This avoids raw log lines appearing on the console before the live UI starts.


def get_partitions_letters() -> List[str]:
    """Get list of available disk partition letters."""
    partitions = psutil.disk_partitions()
    partitions_letters = []
    for partition in partitions:
        partitions_letters.append(partition.mountpoint.replace(":\\", ""))

    return partitions_letters


def generate_category(subject: str) -> Tuple[str, Optional[str]]:
    """
    Classify an email into a (category, sub_category) pair by matching the
    subject against keywords defined in config.yaml.

    Resolution order:
      1. Sub-categories are checked before their parent category so that a
         more-specific match always wins.
      2. The first matching category wins (order follows config.yaml).
      3. Falls back to the 'others' category name when nothing matches.

    Args:
        subject: Email subject string.

    Returns:
        Tuple of (category_name, sub_category_name | None).
    """
    categories = config.get_category_config()
    subject_lower = subject.lower()

    # Resolve the fallback name from config so it stays in sync.
    others_name: str = categories.get('others', {}).get('name', 'Others - أخرى')

    for cat_key, cat_cfg in categories.items():
        if cat_key == 'others' or not isinstance(cat_cfg, dict):
            continue

        cat_name: str = cat_cfg.get('name', cat_key)
        cat_keywords: list = cat_cfg.get('keywords', [])
        sub_categories: dict = cat_cfg.get('sub_categories', {})

        # Check whether the subject matches this top-level category at all.
        cat_matched = any(kw.lower() in subject_lower for kw in cat_keywords)
        if not cat_matched:
            continue

        # Category matched — now check for a more-specific sub-category.
        for sub_key, sub_cfg in sub_categories.items():
            if not isinstance(sub_cfg, dict):
                continue
            sub_keywords: list = sub_cfg.get('keywords', [])
            if any(kw.lower() in subject_lower for kw in sub_keywords):
                sub_name: str = sub_cfg.get('name', sub_key)
                return (cat_name, sub_name)

        # No sub-category matched — check for the special mv3 sub-category
        # (daily_report only: sub_category equals the category name itself).
        mv3_cfg = sub_categories.get('mv3', {})
        if mv3_cfg:
            mv3_keywords: list = mv3_cfg.get('keywords', [])
            if any(kw.lower() in subject_lower for kw in mv3_keywords):
                return (cat_name, cat_name)

        return (cat_name, None)

    return (others_name, None)


def get_outlook_folders() -> Tuple[Any, Any]:
    """
    Connect to Outlook and get inbox and archive folders.
    
    Returns:
        Tuple of (inbox_folder, archive_folder)
    """
    outlook_config = config.get_outlook_config()
    connect = Connection(
        outlook_config['application'], 
        outlook_config['namespace']
    )
    outlook_namespace = connect.get_namespace()
    outlook_folders = Folder(outlook_namespace)
    logging.info("Opening Outlook folders")
    inbox = outlook_folders.get_by_number(
        folder_number=outlook_config['inbox_folder_number']
    )
    archive = outlook_folders.get_by_name(
        root_folder=outlook_config['archive_root_folder'], 
        folder_name=outlook_config['archive_folder_name']
    )

    return inbox, archive


def validate_user_partition(user_partition: str) -> str:
    """
    Validate if the given partition letter exists.
    
    Args:
        user_partition: Partition letter to validate
        
    Returns:
        Partition letter if valid, 0 otherwise
    """
    partitions_letters = get_partitions_letters()
    if user_partition in partitions_letters:
        return user_partition
    return ""


def get_user_partition() -> str:
    """
    Get partition letter from user input or command line arguments.
    
    Returns:
        Valid partition letter
    """
    partitions_letters = get_partitions_letters()
    if len(sys.argv) > 1:
        user_partition = str(sys.argv[1]).upper()
    else:
        user_partition = input(
            "\n👌 Please Enter partition letter to save to: "
        ).upper()

    while not validate_user_partition(user_partition):
        user_partition = input(
            "👌 Please, Enter a valid Partition Letter to save to: "
        ).upper()
        print()

    return user_partition


def get_output_dir(user_partition: str) -> Path:
    """
    Create and return output directory path.
    
    Args:
        user_partition: Partition letter
        
    Returns:
        Path object for output directory
    """
    base_folder = config.get_output_base_folder()
    year_format = config.get_year_format().format(year=date.today().year)
    output_dir = Path(f"{user_partition}:\\{base_folder}\\{year_format}\\")
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"Output directory created: {output_dir}")
    return output_dir


def save_attachments(
    mail: Any, 
    attachments: Any, 
    output_dir: Path, 
    category: str, 
    sub_category: Optional[str], 
    compound: str
) -> None:
    """
    Save attachments from an email.
    
    Args:
        mail: Email message object
        attachments: Attachments collection
        output_dir: Base output directory
        category: Email category
        sub_category: Email sub-category
        compound: Sender compound name
    """
    for attachment in attachments:
        attachment_instance = Attachment(attachment)

        if not attachment_instance.is_ignored() and attachment_instance.accepted_type():
            folder_path = AttachmentPath().attachment_path(
                output_dir=output_dir,
                category=category,
                sub_category=sub_category,
                compound=compound,
                month=attachment_instance.attachment_month(mail),
            )
            file_path = attachment_instance.attachment_folder(folder_path)
            attachment_instance.save_attachment(file_path)
            logging.info(f"Saved attachment: {file_path}")




def get_compound(mail: Mail) -> str:
    """
    Extract and normalize compound name from sender.
    
    Args:
        mail: Mail object
        
    Returns:
        Normalized compound name in uppercase
    """
    return (
        str(mail.mail.sender)
        .lower()
        .replace(" security", "")
        .replace(" buildingsecurity", "")
        .upper()
    )


def process_mail(
    item: Any,
    archive: Any,
    output_dir: Path,
    counter: int,
    unread_flag: str,
    ui: Optional[ProgressUI] = None,
) -> int:
    """Process a single email and return the updated counter."""
    if item.unread and unread_flag != "Y":
        return counter

    mail_message = Mail(item)
    category, sub_category = generate_category(str(mail_message.mail.subject))

    try:
        counter += 1
        logging.info(f"Processing mail {counter} from {mail_message.mail.sender}")

        compound = get_compound(mail_message)

        if compound == "CONDOLENCES":
            mail_message.mark_read()
            mail_message.move_mail(archive)
            return counter

        save_attachments(
            mail_message.mail,
            mail_message.get_mail_attachments(),
            output_dir,
            category,
            sub_category,
            compound,
        )

        mail_message.move_mail(archive)

    except Exception as exc:
        logging.error(f"Error processing mail: {exc}")
        if ui:
            ui.error(f"✗  {exc}")

    return counter


def process_all_mails(
    inbox: Any,
    archive: Any,
    output_dir: Path,
    unread_flag: str,
    ui: Optional[ProgressUI] = None,
) -> int:
    """
    Process all emails in inbox.

    Args:
        inbox:       Inbox folder.
        archive:     Archive folder.
        output_dir:  Output directory.
        unread_flag: 'Y' or 'N' flag.
        ui:          Optional ProgressUI instance for live updates.

    Returns:
        Total processed emails count.
    """
    counter = 0
    items = list(inbox.items)
    total = len(items)
    logging.info(f"Found {total} items in inbox")

    for idx, item in enumerate(items, start=1):
        counter = process_mail(item, archive, output_dir, counter, unread_flag, ui)
        if ui:
            ui.update(current=idx)

    return counter




def main() -> None:
    """Main entry point for the application."""
    # Configure logging immediately — file handler only, no console output.
    # The live UI handles all user-visible feedback from this point on.
    setup_logging(ui=None, file_only=True)

    user_partition = get_user_partition()
    output_dir = get_output_dir(user_partition)
    inbox, archive = get_outlook_folders()

    is_unread = any(item.unread for item in list(inbox.items))
    unread = input("\n👀 Save and Archive unread mails? Y(es)/N(o): ") if is_unread else "N"

    items = list(inbox.items)
    process_total = (
        len(items) if unread.upper() == "Y"
        else sum(1 for it in items if not it.unread)
    )

    # ── Start live UI ──────────────────────────────────────────────────
    ui = ProgressUI(total=process_total if process_total > 0 else None)
    set_ui(ui)
    ui.start()

    try:
        counter = process_all_mails(inbox, archive, output_dir, unread, ui)

        summary = (
            f"🎊  All mails archived and attachments saved  ({counter} processed)"
            if unread.upper() == "Y"
            else f"🎊  All read mails archived and attachments saved  ({counter} processed)"
        )
        ui.complete(summary)
        logging.info(summary)

    finally:
        import time
        time.sleep(2.0)
        ui.stop()
        set_ui(None)

    print()


if __name__ == "__main__":
    main()
