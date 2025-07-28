from win32com.client import Dispatch


class Connection:

    def __init__(self, application, namesapce):
        self.application = application
        self.namespace = namesapce

    def connect(self):
        return Dispatch(self.application)

    def get_namespace(self):
        return self.connect().GetNameSpace(self.namespace)


class Folder:
    def __init__(self, namespace):
        self.namespace = namespace

    def get_default_folder(self, folder_number):
        return self.namespace.GetDefaultFolder(folder_number)

    def get_folder(self, root_folder, folder_name):
        return self.namespace.Folders(root_folder).Folders(folder_name)
