# to correctly display arabic language in terminal
from bidi.algorithm import get_display

from datetime import date
from pathlib import Path


class AttachmentPath:

    def attachment_path(self, output_dir, category, compound, month, sub_category=None):

        if "Tech" in category:
            week_number = (int(date.today().day) - 1) // 7 + 1

        path = f"{output_dir}\\{category}\\{compound}\\{sub_category if sub_category else ''}\\{month}\\{'week '+str(week_number) if "Tech" in category else ''}"
        return path.replace("\\\\", "\\")


class Mail:

    def __init__(self, message):
        self.mail = message

    def get_mail_attachments(self):
        return self.mail.attachments

    def move_mail(self, folder):
        """
        move mail to the given folder
        """

        try:
            self.mail.Move(folder)
            print(f"\t☑️ {get_display(self.mail.subject)}")
        except Exception as e:
            print(f"\n🤯 Faced error while trying to move the email to {folder}: {e}\n")

    def mark_read(self):
        self.mail.Unread = False
        self.mail.Save()

    def is_read(self):
        return not self.mail.unread


class Attachment:

    def __init__(self, attachment):
        self.attachment = attachment

    def is_ignored(self):
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
        return str(self.attachment.filename) in ignored_files

    def accepted_type(self):
        ext = ["docx", "pdf", "xlsx", "pptx"]
        attachment_type = str(self.attachment.filename).split(".")[-1]
        return attachment_type in ext

    def attachment_month(self, item):

        filename = str(self.attachment.filename)
        for letter in filename:
            if not letter.isnumeric() and letter != "-" and letter != ".":
                filename = filename.replace(letter, "")

        init_date = filename.replace(".", "-").strip("-").strip()
        if not init_date[-4:].isnumeric():
            if init_date != 0:
                init_date = init_date[:-2]
        date_list = init_date.split("-")
        month_number = None
        if len(date_list) >= 3:
            month_number = init_date.split("-")[-2] or 0
            month_number = int(month_number)

        if month_number not in [month for month in range(1, 13)]:
            month_number = item.ReceivedTime.month

        dt = date(2000, month_number, 1)
        month_name = dt.strftime("%B")
        return f"{month_number}. {month_name}"

    def attachment_folder(self, path):
        path = Path(path)
        path.mkdir(parents=True, exist_ok=True)
        file_path = str(path.absolute()) + "\\" + str(self.attachment.filename)
        return file_path

    def save_attachment(self, file_path):
        if not self.is_ignored() and self.accepted_type():
            self.attachment.SaveAsFile(file_path)
        elif not self.accepted_type():
            print(f"\n⁉️ Unsupported Attachment Type. '{self.attachment.filename}'\n")
