import datetime
import os
import shutil
import subprocess
import tempfile
import win32com.client
import sys

from Components import OpenFile
from Worker import SearchWorker

shell = win32com.client.Dispatch("WScript.Shell") #access window shell
desktop = shell.SpecialFolders("Desktop") #access windows desktop file(not oneDrive)
documents = shell.SpecialFolders("MyDocuments")
recycle = shell.SpecialFolders("RecycleBin")

from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout,
    QLineEdit, QPushButton, QListWidget, QLabel, QListWidgetItem, QDialog, 
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon


class SFSearch(QWidget):
    
    #initial self
    def __init__(self, switch_page_callback):
        super().__init__()
        
        self.fileData = []
        
        layout = QVBoxLayout() #hold functions
        
        self.label = QLabel("Enter filename (or part of it):")  #add a label to layout
        layout.addWidget(self.label)
        
        # input box
        self.input = QLineEdit() #add search bar
        layout.addWidget(self.input)
        
        
        
        # Buttons
        
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        user = os.path.join(os.path.expanduser("~"))
        program_files     = "C:\\Program Files"
        program_files_x86 = "C:\\Program Files (x86)"
        self.button_layout_1 = QHBoxLayout()

        self.desktopbtn = QPushButton("Desktop")
        self.desktopbtn.clicked.connect(lambda: self.perform_search(desktop))

        self.docBtn = QPushButton("Documents")
        self.docBtn.clicked.connect(lambda: self.perform_search(documents)) 
        
        self.downBtn = QPushButton("Downloads")
        self.downBtn.clicked.connect(lambda: self.perform_search(downloads)) 

        self.PFbtn = QPushButton("Program Files")
        self.PFbtn.clicked.connect(lambda: self.perform_search(program_files))

        self.PFx86btn = QPushButton("Program Files (x86)")
        self.PFx86btn.clicked.connect(lambda: self.perform_search(program_files_x86)) 

        self.tempbtn = QPushButton("%TEMP%")
        self.tempbtn.clicked.connect(lambda: self.perform_search(tempfile.gettempdir()))

        self.userbtn = QPushButton("Users")
        self.userbtn.clicked.connect(lambda: self.perform_search(user))

        self.button_layout_1.addWidget(self.desktopbtn)     
        self.button_layout_1.addWidget(self.docBtn)
        self.button_layout_1.addWidget(self.downBtn)
        self.button_layout_1.addWidget(self.PFbtn)
        self.button_layout_1.addWidget(self.PFx86btn)
        self.button_layout_1.addWidget(self.tempbtn)
        self.button_layout_1.addWidget(self.userbtn) 

        self.searchAllbtn = QPushButton("Search Entire Drive") #Add button 
        self.searchAllbtn.clicked.connect(lambda: self.perform_search("C:\\")) #add function for when button is clicked 
        layout.addLayout(self.button_layout_1)
        layout.addWidget(self.searchAllbtn)

        # results list
        self.results = QListWidget() #Add results
        layout.addWidget(self.results)

        self.setLayout(layout)
        
        
        
    def perform_search(self, base_path):
        keyword = self.input.text().strip()
        self.results.clear()

        if not keyword:
            self.results.addItem("Please enter a search term")
            return
        
        self.worker = SearchWorker(keyword, base_path)
        self.worker.results_found.connect(self.display_results)
        self.worker.progress.connect(self.update_status)
        self.worker.start()
        
        
        self.desktopbtn.setEnabled(False)
        self.docBtn.setEnabled(False)
        self.downBtn.setEnabled(False)
        self.tempbtn.setEnabled(False)
        self.PFbtn.setEnabled(False)
        self.PFx86btn.setEnabled(False)
        self.userbtn.setEnabled(False)
        self.searchAllbtn.setEnabled(False)

        self.results.itemDoubleClicked.connect(self.open_file)   
         
    def display_results(self, matches):
        self.desktopbtn.setEnabled(True)
        self.docBtn.setEnabled(True)
        self.downBtn.setEnabled(True)
        self.tempbtn.setEnabled(True)
        self.PFbtn.setEnabled(True)
        self.PFx86btn.setEnabled(True)
        self.userbtn.setEnabled(True)
        self.searchAllbtn.setEnabled(True)
        if matches:
            for m in matches:
                self.results.addItem(m)
        else: 
            self.results.addItem("No matching file found")
    
    def update_status(self, message): self.label.setText(message) # reuse label as a "status"
    def open_file(self, item):
        metadata = item.data(Qt.UserRole)
        new_window = OpenFile(metadata)
        new_window.show()
        if not hasattr(self, "open_windows"):
            self.open_windows = []
        self.open_windows.append(new_window)  



