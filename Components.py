
import os
import shutil
import subprocess
import tempfile
import win32com.client
import sys


shell = win32com.client.Dispatch("WScript.Shell") #access window shell
desktop = shell.SpecialFolders("Desktop") #access windows desktop file(not oneDrive)
documents = shell.SpecialFolders("MyDocuments")
recycle = shell.SpecialFolders("RecycleBin")

from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout,
    QLineEdit, QPushButton, QListWidget, QLabel, QListWidgetItem, QDialog, 
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt





class OpenFile(QWidget):
    def __init__(self, file):
        super().__init__()
        self.file = file
        self.setWindowTitle("Open File") #set title
        self.setFixedSize(300,300)
        layout = QVBoxLayout() 
        self.list = QListWidget()
        self.list.addItem(self.file["name"])
        layout.addWidget(self.list)
        
        self.openbtn = QPushButton("Open")
        self.openbtn.clicked.connect(self.open_file)
        
        
        self.openInDrivebtn = QPushButton("Open in File Location")
        self.openInDrivebtn.clicked.connect(self.open_in_folder)

        self.copybtn = QPushButton("Copy to Desktop")
        self.copybtn.clicked.connect(self.copy_to)

        
        layout.addWidget(self.openbtn)
        layout.addWidget(self.openInDrivebtn)
        layout.addWidget(self.copybtn)
        self.setLayout(layout)
        
        
    def open_file(self):
        
        if os.path.exists(self.file["path"]):
                os.startfile(self.file["path"])
        else:
                print("file not opening")
        
    def open_in_folder(self):
            if os.path.exists(self.file["path"]):
                subprocess.Popen(["explorer", "/select,", self.file["path"]]) #open in folder
            else:
                print("file not opening")
                
    def copy_to(self):
        if os.path.exists(self.file["path"]):
            try:
                shutil.copy2(self.file["path"], desktop)
                new_window = UpdateUser("Copy to Folder", "Copied to Desktop Successfully")
                new_window.show()
                if not hasattr(self, "open_windows"):
                    self.open_windows = []
                self.open_windows.append(new_window)  

            except Exception as e:
                print(f"Error: {e}")
                new_window = UpdateUser("Copy to Folder", "Error: Could not copy to desktop :( ")

        






class UpdateUser(QDialog):
    def __init__(self, title, update):
        super().__init__()
        self.setWindowTitle(title) 
        layout = QVBoxLayout() 
        self.label  = QLabel(f"{update}")
        
        layout.addWidget(self.label)
        self.setLayout(layout)





