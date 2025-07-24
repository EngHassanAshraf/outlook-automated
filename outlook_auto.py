# This is a file to automate outlook mails' attachments saving

# after saving attachments => move the mail to deleted folder number 3

# import asyncio
# from googletrans import Translator

import os
from pathlib import Path
from datetime import date
from win32com.client import Dispatch


# async def translate(text):
#     async with Translator() as translator:
#         translated = await translator.translate(text, dest="ar")
#     return translated


def create_folder(sender, report_type):
    month_name_en = date.today().strftime("%B")
    # month_name_ar = asyncio.run(translate(month_name_en))
    month_number = date.today().month
    week_number = (int(date.today().day) - 1) // 7 + 1

    output_dir = Path("./MV-2025/")
    output_dir.mkdir(parents=True, exist_ok=True)

    if report_type == "Tech Report":
        attachment_folder = (
            output_dir
            / report_type
            / sender
            / f"{month_number}. {month_name_en}"
            / f"week {week_number}"
        )
    elif report_type == "Accommodation" or report_type == "Dayra":
        attachment_folder = (
            output_dir
            / "Daily Report"
            / sender
            / report_type
            / f"{month_number}. {month_name_en}"
        )
    elif report_type == "Daily Report" and (
        "sokhna" in sender.lower() or "sahel" in sender.lower()
    ):
        attachment_folder = (
            output_dir
            / report_type
            / sender
            / report_type
            / f"{month_number}. {month_name_en}"
        )
    else:
        attachment_folder = (
            output_dir / report_type / sender / f"{month_number}. {month_name_en}"
        )

    attachment_folder.mkdir(parents=True, exist_ok=True)
    return attachment_folder.absolute()


def save_attachments(item, report_type, ignored_files):
    sender = str(item.sender).lower().replace(" security", "").upper()
    attachment_folder = create_folder(sender, report_type)

    for attachment in item.attachments:
        attachment_type = str(attachment.filename).split(".")[-1]
        if (
            # attachment_type == "docx"
            # or attachment_type == "pdf"
            # or attachment_type == "xlsx"
            str(attachment.filename)
            not in ignored_files
        ):
            final_path = str(attachment_folder) + "\\" + str(attachment)
            attachment.SaveAsFile(final_path)


def outlook_folder(folder_name):
    outlook_connection = Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook_connection.Folders(
        "Security.Information@mountainview-eg.com"
    ).Folders(folder_name)
    return folder


def main():
    inbox = outlook_folder("Inbox")
    archive = outlook_folder("Archive")
    ignored_files = [
        "EmailSignature-International_N_374acb21-a63f-4e28-ac6f-11c4b255b559.jpg",
        "fb-resized_1_08ab0129-1fbf-4e09-9caf-52e1f3fb8718.png",
        "insta-resized_2040ce31-0940-4a34-a926-c5194e1f7f3c.png",
        "linkedin-resized_7fce048f-c49e-4e77-8e1d-bbfbcd6aaaf1.png",
        "image001.png",
        "image002.png",
        "image003.png",
        "image004.png",
        "image005.png",
        "image006.png",
        "image007.png",
        "image008.png",
        "image009.png",
        "image010.png",
        "image011.png",
        "image012.png",
        "image013.png",
        "image014.png",
        "image015.png",
        "image016.png",
        "image017.png",
        "image018.png",
        "image019.png",
        "image020.png",
        "image021.png",
        "image022.png",
        "image023.png",
        "image024.png",
        "image025.png",
        "image026.png",
        "image027.png",
        "image028.png",
        "image029.png",
        "image030.png",
        "image031.png",
        "image032.png",
        "image033.png",
        "image034.png",
        "image035.png",
        "image036.png",
        "image037.png",
        "image038.png",
        "image039.png",
        "image001.jpg",
        "image002.jpg",
        "image003.jpg",
        "image004.jpg",
        "image005.jpg",
        "image006.jpg",
        "image007.jpg",
        "image008.jpg",
        "image009.jpg",
        "image010.jpg",
        "image011.jpg",
        "image012.jpg",
        "image013.jpg",
        "image014.jpg",
        "image015.jpg",
        "image016.jpg",
        "image017.jpg",
        "image018.jpg",
        "image019.jpg",
        "image020.jpg",
        "image021.jpg",
        "image022.jpg",
        "image023.jpg",
        "image024.jpg",
        "image025.jpg",
        "image026.jpg",
        "image027.jpg",
        "image028.jpg",
        "image029.jpg",
        "image030.jpg",
        "image031.jpg",
        "image032.jpg",
        "image033.jpg",
        "image034.jpg",
        "image035.jpg",
        "image036.jpg",
        "image037.jpg",
        "image038.jpg",
        "image039.jpg",
    ]
    for item in inbox.Items:
        if (
            "الحالة الفنية" in str(item.subject)
            or "weekly" in str(item.subject)
            or "technical" in str(item.subject).lower()
        ):
            save_attachments(item, "Tech Report", ignored_files)
        elif (
            "تواجد الملاك" in str(item.subject)
            or "mv-nc accommodation" in str(item.subject).lower()
        ):
            save_attachments(item, "Accommodation", ignored_files)

        elif "تواجد الزائرين" in str(item.subject):
            save_attachments(item, "Dayra", ignored_files)

        elif (
            "اليومي" in str(item.subject)
            or "اليومى" in str(item.subject)
            or "الحالة الأمنية" in str(item.subject)
            or "daily" in str(item.subject).lower()
        ):
            save_attachments(item, "Daily Report", ignored_files)

        elif (
            "تعيين" in str(item.subject)
            or "تعين" in str(item.subject)
            or "وثقية" in str(item.subject)
        ):
            save_attachments(item, "أوراق الموظفين", ignored_files)
        else:
            save_attachments(item, "Others", ignored_files)

        item.Move(archive)


if __name__ == "__main__":
    main()
