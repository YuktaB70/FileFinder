import datetime
import os
import shutil
import subprocess
import tempfile
import win32com.client
import sys
import glob

from Components import OpenFile
from Worker import SearchWorker

shell = win32com.client.Dispatch("WScript.Shell") #access window shell
desktop = shell.SpecialFolders("Desktop") #access windows desktop file(not oneDrive)
documents = shell.SpecialFolders("MyDocuments")
recycle = shell.SpecialFolders("RecycleBin")

from PyQt5.QtWidgets import *
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon



class SFDashboard(QWidget):
    def __init__(self, switch_page_callback):
        super().__init__()
        self.setStyleSheet("""
            QLabel {
                position: absolute;
                font-size: 20px;
                margin-top: 10px;
                margin-bottom: 10px;

            }
                                    
            QProgressBar {
                border: none;
                border-radius: 15px;
                width: 20px;
                background: #ccc;
                text-align: center;
            }

            QProgressBar::chunk {
                background-color: #D77D28;
                border-radius: 15px;
                margin: 0px;
            }
           
            QListWidget {
                background: #f5f5f7;
                border-radius: 12px;
                padding: 10px;
                border: none;
                font-size: 15px;
            }

            QListWidget::item {
                background: #ffffff;
                margin: 6px;
                padding: 12px;
                border-radius: 10px;
                color: #333;
            }

            QListWidget::item:selected {
                background: #d0e2ff;
                color: black;
            }
        """)

        layout = QVBoxLayout() 
        self.label = QLabel("Dashboard")
        layout.addWidget(self.label)
        storage = self.get_storage()
        
        progress = (storage["used"] / storage["total"]) * 100  
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(int(progress))
        self.bar.setFixedHeight(60)
        self.bar.setTextVisible(True)
        layout.addWidget(self.bar)
        storage_label = QHBoxLayout()
        used_label = QLabel(f"{storage["used"]} GB Used")
        free_label = QLabel(f"{storage["free"]} GB Free")
        storage_label.addWidget(used_label)
        storage_label.addStretch()
        storage_label.addWidget(free_label)
        layout.addLayout(storage_label)

        recent_label = QLabel("Recent Files")
        
        self.recent_file = QListWidget()

        self.worker = SearchWorker("null", "C:\\", "find recent files")
        self.worker.results_found.connect(self.display_results)
        self.worker.progress.connect(self.update_status)
        self.worker.start()

        self.setLayout(layout)
        layout.addWidget(recent_label)
        layout.addWidget(self.recent_file)
        

    
    
    def get_storage(self):
        total_disk, used, free = shutil.disk_usage("C:\\")
        total_disk = round(total_disk/(1024**3), 2)
        used = round(used/(1024**3), 2)
        free = round(free/(1024**3), 2)

        
        return {
            "total": total_disk,
            "used": used,
            "free": free
        }
        
    def display_results(self, matches):
        self.recent_file.setSortingEnabled(False)
        self.recent_file.clear() 
        if matches:
            for m in matches:
                self.recent_file.addItem(m)
        else: 
            self.recent_file.addItem("No recent file found")

    def update_status(self, message): self.label.setText(message) 