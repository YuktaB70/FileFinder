import datetime
import os
import shutil
import subprocess
import tempfile
import win32com.client
import sys


shell = win32com.client.Dispatch("WScript.Shell") #access window shell
desktop = shell.SpecialFolders("Desktop") #access windows desktop file(not oneDrive)
documents = shell.SpecialFolders("MyDocuments")

from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout,
    QLineEdit, QPushButton, QListWidget, QLabel, QListWidgetItem
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
#QApplicaiton -> main app manager
#QWidget -> the base class for windows
#QVBoxLayout -> arranges widgets vertically(in column)
#QlineEdit -> a textbox 
#QPushButton -> creates a clickable button
#QlistWidget -> creates a list widget
#QLabel -> label widget


#Main window app. Built on top of QWidget
class FileFinderApp(QWidget):
    
    #initial self
    def __init__(self):
        super().__init__()
        
        self.fileData = []
        
        self.setWindowTitle("File Finder") #set title
        self.setFixedSize(700,400) #set size of window box
        layout = QVBoxLayout() #Create a vertical layout
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

        self.buttons = QHBoxLayout()
       
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

        
        self.buttons.addWidget(self.desktopbtn)     
        self.buttons.addWidget(self.docBtn)
        self.buttons.addWidget(self.downBtn)
        self.buttons.addWidget(self.PFbtn)
        self.buttons.addWidget(self.PFx86btn)
        self.buttons.addWidget(self.tempbtn)
        self.buttons.addWidget(self.userbtn) 
        self.button = QPushButton("Search Entire Drive") #Add button 
        self.button.clicked.connect(lambda: self.perform_search("C:\\")) #add function for when button is clicked 
        layout.addLayout(self.buttons)
        layout.addWidget(self.button)

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
        
        self.results.itemDoubleClicked.connect(self.open_file)   
         
    def display_results(self, matches):
        if matches:
            for m in matches:
                self.results.addItem(m)
        else: 
            self.results.addItem("No matching file found")
    
    def update_status(self, message): self.label.setText(message) # reuse label as a "status"
    def open_file(self, item):
        new_window = OpenFile(item)
        new_window.show()
        if not hasattr(self, "open_windows"):
            self.open_windows = []
        self.open_windows.append(new_window)  












class OpenFile(QWidget):
    def __init__(self, file):
        super().__init__()
        self.file = file
        self.setWindowTitle("Open File") #set title
        self.setFixedSize(300,300)
        layout = QVBoxLayout() 
        self.list = QListWidget()
        self.list.addItem(self.file.text())
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
            if os.path.exists(self.file.text()):
                os.startfile(self.file.text())
            else:
                print("file not opening")
        
    def open_in_folder(self):
            if os.path.exists(self.file.text()):
                subprocess.Popen(["explorer", "/select,", self.file.text()]) #open in folder
            else:
                print("file not opening")
                
    def copy_to(self):
        if os.path.exists(self.file):
            try:
                shutil.copy2(self.file, desktop)
            except Exception as e:
                print(f"Error: {e}")
        










class SearchWorker(QThread):
    results_found = pyqtSignal(list) #this will send the list back
    progress = pyqtSignal(str) #this will send a status back
    
    def __init__(self, keyword, base_path):
        super().__init__()
        self.keyword = keyword.lower()
        self.base_path = base_path
    
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
    
    def run(self):
        results = self.search_file_by_name(self.keyword, self.base_path)
        self.results_found.emit(results)
    

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileFinderApp()
    window.show()
    sys.exit(app.exec_())
