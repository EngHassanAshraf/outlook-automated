from win32com.client import Dispatch


class Connection:

    def __init__(self, application, namesapce):
        self.application = application
        self.namespace = namesapce

    def connect(self):
        try:
            print("Connecting\n")
            d = Dispatch(self.application)
            print("Connected Successfully")
            return d
        except Exception as e:
            print(f"Faced an error while connecting: {e}")

    def get_namespace(self):
        try:
            return self.connect().GetNameSpace(self.namespace)
        except Exception as e:
            print(f"Faced an error while getting the namespace: {e}")


class Folder:
    def __init__(self, namespace):
        self.namespace = namespace

    def get_default_folder(self, folder_number):
        try:
            print(f"Folder {folder_number} opened")
            return self.namespace.GetDefaultFolder(folder_number)
        except Exception as e:
            print("Faced an error while openning the folder: {e}")

    def get_folder(self, root_folder, folder_name):
        try:
            print(f"Folder {root_folder}.{folder_name} opened")
            return self.namespace.Folders(root_folder).Folders(folder_name)
        except Exception as e:
            print("Faced an error while openning the folder: {e}")
