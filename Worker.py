import datetime
import os
import win32com.client


shell = win32com.client.Dispatch("WScript.Shell") #access window shell
desktop = shell.SpecialFolders("Desktop") #access windows desktop file(not oneDrive)
documents = shell.SpecialFolders("MyDocuments")
downloads = os.path.join(os.path.expanduser("~"), "Downloads")
EXTENSIONS = [".png",".pdf",".docx",".jpeg",".dwg",".txt",".pptx",".ppt",".xls"]

from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout,
    QLineEdit, QPushButton, QListWidget, QLabel, QListWidgetItem, QDialog, 
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt



class SearchWorker(QThread):
    results_found = pyqtSignal(list) #this will send the list back
    progress = pyqtSignal(str) #this will send a status back
    
    def __init__(self, keyword, base_path, mode):
        super().__init__()
        self.keyword = keyword.lower()
        self.base_path = base_path
        self.mode = mode
    
    def search_file_by_name(self, keyword, base_path):
        matches = []
        keyword = keyword.lower()
        for root, dirs, files in os.walk(base_path, topdown=True, followlinks=False):
            for f in files:
                if keyword in f.lower():
                    file_path = os.path.join(root, f)
                    file_name = os.path.basename(file_path)
                    
                    try: 
                        mtime = os.path.getmtime(file_path)
                        date = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
                    except Exception:
                        date = "N/A"
                    
                    
                    metadata = {
                        "name": file_name,
                        "path": file_path,
                        "date": date 
                    }
                    item = QListWidgetItem(f"{file_name} ---------------------- Last Modified: {date}")
                    item.setData(Qt.UserRole, metadata)
                    
                    
                    
                    matches.append(item)
                    
                    
            self.progress.emit(f"Scanning {root}...")

        return matches

    def search_recent(self, base_path):
        recent = []
        
        for root, dirs, files in os.walk(base_path, topdown=True, followlinks=True):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in EXTENSIONS:
                    full_path = os.path.join(root, f)

                    try:
                        mtime = os.path.getmtime(full_path)
                    except PermissionError:
                        continue

                    date = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")

                    metadata = {
                    "name": f,
                    "path": full_path,
                    "date": date
                    }

                    item = QListWidgetItem(f"{f} ---------------------------------------------------------- Last Modified: {date}")
                    item.setData(Qt.UserRole, metadata)

                    recent.append((mtime, item))

        recent.sort(key=lambda x: x[0], reverse=True)

        return [item for _, item in recent[0:10]]

    def run(self):
        if self.mode == "search by keyword":
            results = self.search_file_by_name(self.keyword, self.base_path)
            self.results_found.emit(results)
        if self.mode == "find recent files":
            results = self.search_recent(downloads)
            self.results_found.emit(results)
