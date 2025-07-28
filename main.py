from pathlib import Path
from datetime import date

from application import Connection, Folder
from content import Message, Attachment, AttachmentPath


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
        if "mv3" in subject.lower() or "السخنة" in subject:
            sub_category = category

    elif "تعيين" in subject or "تعين" in subject or "وثقية" in subject:
        category = "Staff - الموظفين"
    else:
        category = "Others - أخرى"

    return (category, sub_category)


def main():
    connect = Connection("Outlook.Application", "MAPI")
    outlook_namespace = connect.get_namespace()
    outlook_folders = Folder(outlook_namespace)

    inbox = outlook_folders.get_default_folder(folder_number=6)
    archive = outlook_folders.get_folder(root_folder="Archives", folder_name="Archive")

    output_dir = Path(f"D:\\MV\\MV-{date.today().year}\\")
    output_dir.mkdir(parents=True, exist_ok=True)

    for item in list(archive.items):

        category, sub_category = generate_category(str(item.subject))
        compound = (
            str(item.sender)
            .lower()
            .replace(" security", "")
            .replace(" buildingsecurity", "")
            .upper()
        )

        try:
            message = Message(item)
            if compound == "CONDOLENCES":
                # message.move_message(folder=archive)
                continue

            attachments = message.get_message_attachments()
            for attachment in attachments:
                attachment_instance = Attachment(attachment)
                if (
                    not attachment_instance.is_ignored()
                    and attachment_instance.accepted_type()
                ):

                    folder_path = AttachmentPath().attachment_path(
                        output_dir=output_dir,
                        category=category,
                        sub_category=sub_category,
                        compound=compound,
                        month=attachment_instance.attachment_month(item),
                    )
                    file_path = attachment_instance.attachment_folder(folder_path)
                    attachment_instance.save_attachment(file_path=file_path)
            # message.move_message(folder=archive)
        except Exception as e:
            print(e)


if __name__ == "__main__":
    main()
