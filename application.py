import logging
from typing import Any

from win32com.client import Dispatch


class Connection:
    """Handles connection to Outlook application via COM."""
    
    def __init__(self, application: str, namespace: str):
        """
        Initialize Outlook connection.
        
        Args:
            application: Outlook application name (e.g., "Outlook.Application")
            namespace: Namespace type (e.g., "MAPI")
        """
        self.application = application
        self.namespace = namespace

    def connect(self) -> Any:
        """
        Connect to Outlook application.
        
        Returns:
            Dispatched Outlook application object, or None on failure
        """
        try:
            logging.info("Connecting to Outlook...")
            dispatched = Dispatch(self.application)
            if dispatched:
                logging.info("Connected to Outlook successfully")
            return dispatched
        except Exception as e:
            logging.error(f"Error connecting to Outlook: {e}")
            return None

    def get_namespace(self) -> Any:
        """
        Get Outlook namespace.
        
        Returns:
            Outlook namespace object, or None on failure
        """
        try:
            outlook_app = self.connect()
            if outlook_app:
                return outlook_app.GetNameSpace(self.namespace)
            return None
        except Exception as e:
            logging.error(f"Error getting namespace: {e}")
            return None


class Folder:
    """Handles Outlook folder operations."""
    
    _default_folders = {"6": "Inbox"}

    def __init__(self, namespace: Any):
        """
        Initialize Folder manager.
        
        Args:
            namespace: Outlook namespace object
        """
        self.namespace = namespace

    def get_by_number(self, folder_number: int) -> Any:
        """
        Get folder by its default number.
        
        Args:
            folder_number: Outlook default folder number (e.g., 6 for Inbox)
            
        Returns:
            Folder object, or None on failure
        """
        try:
            folder_name = self._default_folders.get(str(folder_number), "Unknown")
            logging.info(f"Opening folder: {folder_name} (#{folder_number})")
            return self.namespace.GetDefaultFolder(folder_number)
        except Exception as e:
            logging.error(f"Error opening folder #{folder_number}: {e}")
            return None

    def get_by_name(self, root_folder: str, folder_name: str) -> Any:
        """
        Get folder by name from a specific root folder.
        
        Args:
            root_folder: Root folder name (e.g., "Archives")
            folder_name: Folder name within root
            
        Returns:
            Folder object, or None on failure
        """
        try:
            logging.info(f"Opening folder: {root_folder}\\{folder_name}")
            return self.namespace.Folders(root_folder).Folders(folder_name)
        except Exception as e:
            logging.error(f"Error opening folder {root_folder}\\{folder_name}: {e}")
            return None
