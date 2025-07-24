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


def outlook_folder(folder_number):
    outlook_connection = Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook_connection.GetDefaultFolder(folder_number)
    return folder


def main():
    inbox = outlook_folder(6)
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
        "image0010.png",
        "image0011.png",
        "image0012.png",
        "image0013.png",
        "image0014.png",
        "image0015.png",
        "image0016.png",
        "image0017.png",
        "image0018.png",
        "image0019.png",
        "image0020.png",
        "image0021.png",
        "image0022.png",
        "image0023.png",
        "image0024.png",
        "image0025.png",
        "image0026.png",
        "image0027.png",
        "image0028.png",
        "image0029.png",
        "image0030.png",
        "image0031.png",
        "image0032.png",
        "image0033.png",
        "image0034.png",
        "image0035.png",
        "image0036.png",
        "image0037.png",
        "image0038.png",
        "image0039.png",
        "image001.jpg",
        "image002.jpg",
        "image003.jpg",
        "image004.jpg",
        "image005.jpg",
        "image006.jpg",
        "image007.jpg",
        "image008.jpg",
        "image009.jpg",
        "image0010.jpg",
        "image0011.jpg",
        "image0012.jpg",
        "image0013.jpg",
        "image0014.jpg",
        "image0015.jpg",
        "image0016.jpg",
        "image0017.jpg",
        "image0018.jpg",
        "image0019.jpg",
        "image0020.jpg",
        "image0021.jpg",
        "image0022.jpg",
        "image0023.jpg",
        "image0024.jpg",
        "image0025.jpg",
        "image0026.jpg",
        "image0027.jpg",
        "image0028.jpg",
        "image0029.jpg",
        "image0030.jpg",
        "image0031.jpg",
        "image0032.jpg",
        "image0033.jpg",
        "image0034.jpg",
        "image0035.jpg",
        "image0036.jpg",
        "image0037.jpg",
        "image0038.jpg",
        "image0039.jpg",
    ]

    for item in inbox.Items:
        if (
            "الحالة الفنية" in str(item.subject)
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
            or "daily" in str(item.subject).lower()
        ):
            save_attachments(item, "Daily Report", ignored_files)

        elif "تعيين جديد" in str(item.subject):
            save_attachments(item, "أوراق الموظفين", ignored_files)

        else:
            save_attachments(item, "Others", ignored_files)


if __name__ == "__main__":
    main()
