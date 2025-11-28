import datetime
import os
import shutil
import subprocess
import tempfile
import win32com.client
import sys

from Components import OpenFile
from Worker import SearchWorker
from SFSearch import SFSearch
from SFDashboard import SFDashboard
shell = win32com.client.Dispatch("WScript.Shell") #access window shell
desktop = shell.SpecialFolders("Desktop") #access windows desktop file(not oneDrive)
documents = shell.SpecialFolders("MyDocuments")
recycle = shell.SpecialFolders("RecycleBin")

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
#QApplicaiton -> main app manager
#QWidget -> the base class for windows
#QVBoxLayout -> arranges widgets vertically(in column)
#QlineEdit -> a textbox 
#QPushButton -> creates a clickable button
#QlistWidget -> creates a list widget
#QLabel -> label widget


#Main window app. Built on top of QWidget
class FileFinderApp(QMainWindow):
    
    #initial self
    def __init__(self):
        super().__init__()
        self.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                background: white;
                padding: 10px 15px;
                color: black;
            }
            QLineEdit {
                border-radius: 15px;
                background: white;
                border: none;
                padding: 3px 10px;
                height: 40px;
                
            }
            QListWidget {
                border-radius: 15px;
                background: white;
                border: none;
                margin: 5px;

            }
            

        """
            
        )
        self.setWindowTitle("SystemFlow") #set title
        self.setWindowIcon(QIcon('Logo.png'))
        self.setFixedSize(1250,700) #set size of window box
        
        central = QWidget()
        self.setCentralWidget(central)
        self.container = QHBoxLayout()
        self.btnContainer = QVBoxLayout()
        self.container.addLayout(self.btnContainer)
        central.setLayout(self.container)
        

        self.searchBtn = QPushButton("Search")
        

        self.dashboardBtn = QPushButton("DashBoard")
        self.searchBtn.clicked.connect(lambda: self.switch_page("search"))
        self.dashboardBtn.clicked.connect(lambda: self.switch_page("dashboard"))

        self.btnContainer.addWidget(self.dashboardBtn)
        self.btnContainer.addWidget(self.searchBtn)

        self.btnContainer.addStretch()
        self.btnContainer.addSpacing(1)
        

        self.stack = QStackedWidget()
        self.container.addWidget(self.stack)
        self.pages = {
            "dashboard": SFDashboard(self.switch_page),
            "search": SFSearch(self.switch_page)

        }
        
        for page in self.pages.values():
            self.stack.addWidget(page)
        
        self.switch_page("dashboard")
        
    def switch_page(self, name):
        self.stack.setCurrentWidget(self.pages[name])
  








if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = FileFinderApp()
    window.show()
    sys.exit(app.exec_())
