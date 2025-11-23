from pathlib import Path
from datetime import date

from application import Connection, Folder
from message import Mail, Attachment, AttachmentPath

import psutil
import sys


def get_partitions_letters():
    partitions = psutil.disk_partitions()
    partitions_letters = []
    for partition in partitions:
        partitions_letters.append(partition.mountpoint.replace(":\\", ""))

    return partitions_letters


def generate_category(subject):
    category = None
    sub_category = None

    if "الحالة الفنية" in subject or "technical" in subject.lower():
        category = "Technical Report - التقرير الفني"

    elif "تواجد الملاك" in subject or "mv-nc accommodation" in subject.lower():
        category = "Daily Report - التقرير اليومي"
        sub_category = "Accommodation - تقرير الإقامة"

    elif "تواجد الزائرين" in subject:
        category = "Daily Report - التقرير اليومي"
        sub_category = "Dayra - تقرير دايرة"

    elif "الحضور" in subject or "الإنصراف" in subject:
        category = "Daily Report - التقرير اليومي"
        sub_category = "الحضور والإنصراف"

    elif (
        "اليومي" in subject
        or "اليومى" in subject
        or "الحالة الأمنية" in subject
        or "daily" in subject.lower()
    ):
        category = "Daily Report - التقرير اليومي"
        if "لموقع Mv3" in subject or "السخنة" in subject:
            sub_category = category

    elif (
        "الشركة" in subject
        or "شركة" in subject
        or "الشركه" in subject
        or "شركه" in subject
        or "الشركات" in subject
        or "شركات" in subject
        or "الخارجية" in subject
        or "خارجية" in subject
        or "الخارجيه" in subject
        or "خارجيه" in subject
        or "فاتوره" in subject
        or "فاتورة" in subject
        or "فواتير" in subject
    ):
        category = "الشركات الخارجية"
    elif "التشغيلة" in subject:
        category = "التشغيلات الاسبوعية"
    elif "تعيين" in subject or "تعين" in subject or "وثقية" in subject:
        category = "Staff - الموظفين"
    else:
        category = "Others - أخرى"

    return (category, sub_category)


def get_outlook_folders():
    connect = Connection("Outlook.Application", "MAPI")
    outlook_namespace = connect.get_namespace()
    outlook_folders = Folder(outlook_namespace)

    inbox = outlook_folders.get_by_number(folder_number=6)
    archive = outlook_folders.get_by_name(root_folder="Archives", folder_name="Archive")

    return inbox, archive


def validate_user_partition(user_parition):
    partitions_letters = get_partitions_letters()
    if user_parition in partitions_letters:
        return user_parition
    return 0


def get_user_partiton():
    partitions_letters = get_partitions_letters()
    if len(sys.argv) > 1:
        user_partition = str(sys.argv[1]).upper()
    else:
        user_partition = input(
            "\n👌 Please Enter partition letter to save to: "
        ).upper()

    while not validate_user_partition(user_partition):
        print(f"\n😴 Available Partitions are {partitions_letters}")
        user_partition = input(
            "👌 Please, Enter a valid Partition Letter to save to: "
        ).upper()
        print()

    return user_partition


def get_output_dir(user_partition):
    output_dir = Path(f"{user_partition}:\\MV\\MV-{date.today().year}\\")
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def save_attachments(mail, attachments, output_dir, category, sub_category, compound):

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


def extracting_msg(inbox, unread_flag):
    if unread_flag:
        items = inbox.items
        if len(items):
            print(f"\n\t🔁 Start Extracting all {len(items)} Mails")
        else:
            print(f"\n\t🤔 {len(items)} Mails")

    else:
        items = []
        items = [item for item in inbox.items if not item.unread]
        if items:
            print(f"\n\t🔁 Start Extracting {len(items)} Read Mails")
        else:
            print(f"\n\t🤔 {len(items)} Read Mails")


def get_compound(mail):
    return (
        str(mail.mail.sender)
        .lower()
        .replace(" security", "")
        .replace(" buildingsecurity", "")
        .upper()
    )


def process_mail(item, archive, output_dir, counter, unread_flag):
    if item.unread and unread_flag != "Y":
        return counter

    mail_message = Mail(item)
    category, sub_category = generate_category(str(mail_message.mail.subject))
    try:
        counter += 1
        print(f"\nMail {counter} from {mail_message.mail.sender}")

        compound = get_compound(mail_message)

        if compound == "CONDOLENCES":
            mail_message.mark_read()
            mail_message.move_mail(archive)
            return counter

        attachments = mail_message.get_mail_attachments()

        save_attachments(
            mail_message.mail,
            attachments,
            output_dir,
            category,
            sub_category,
            compound,
        )

        mail_message.move_mail(archive)
    except Exception as e:
        print(f"\n🤯 {e}\n")

    return counter


def process_all_mails(inbox, archive, output_dir, unread_flag):
    counter = 0
    items = list(inbox.items)  # type: ignore

    for item in items:  # type: ignore
        counter = process_mail(item, archive, output_dir, counter, unread_flag)
    return counter


def print_summary(unread_flag):
    if unread_flag == "Y":
        print("\n🎊 All mails moved to Archive folder and its attachments saved\n")
    else:
        print("\n🎊 All read mails moved to Archive folder and its attachments saved\n")


def main():
    user_partition = get_user_partiton()
    output_dir = get_output_dir(user_partition)
    inbox, archive = get_outlook_folders()
    unread = input("\n👀 Save and Archive unread mails? Y(es)/N(o): ")

    extracting_msg(inbox, unread)
    process_all_mails(inbox, archive, output_dir, unread)
    print_summary(unread)


if __name__ == "__main__":
    main()
