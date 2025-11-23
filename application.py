from win32com.client import Dispatch


class Connection:
    def __init__(self, application, namesapce):
        self.application = application
        self.namespace = namesapce

    def connect(self):
        try:
            print("\n🤌 Connecting To Outlook...")
            dispateched = Dispatch(self.application)
            if dispateched:
                print("✅ Connected Successfully")
            return dispateched
        except Exception as e:
            print(f"\n🤯 Faced an error while connecting: {e}\n")

    def get_namespace(self):
        try:
            return self.connect().GetNameSpace(self.namespace)
        except Exception as e:
            print(f"\n🤯 Faced an error while getting the namespace: {e}\n")


class Folder:
    _default_folders = {"6": "Inbox"}

    def __init__(self, namespace):
        self.namespace = namespace

    def get_by_number(self, folder_number):
        try:
            print(f"😉 Folder {self._default_folders[str(folder_number)]} opened")
            return self.namespace.GetDefaultFolder(folder_number)
        except Exception as e:
            print("\n🤯 Faced an error while openning the folder: {e}\n")

    def get_by_name(self, root_folder, folder_name):
        try:
            print(f"😉 Folder {root_folder}.{folder_name} opened")
            return self.namespace.Folders(root_folder).Folders(folder_name)
        except Exception as e:
            print("\n🤯 Faced an error while openning the folder: {e}\n")
