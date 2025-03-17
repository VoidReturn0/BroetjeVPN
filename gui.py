import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QHBoxLayout, QVBoxLayout, QFileSystemModel, QTreeView, QTabWidget,
    QFormLayout, QSplitter, QCheckBox, QTextEdit, QProgressBar
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QDir, pyqtSignal, QTimer, QThread

#########################################
# Worker thread to run the connection command
#########################################
class ConnectionWorker(QThread):
    finished_signal = pyqtSignal(object)
    
    def __init__(self, command, timeout):
        super().__init__()
        self.command = command
        self.timeout = timeout
        
    def run(self):
        try:
            result = subprocess.run(self.command, capture_output=True, text=True, timeout=self.timeout)
            self.finished_signal.emit(result)
        except subprocess.TimeoutExpired as e:
            error_msg = f"Command timed out after {self.timeout} seconds. Please check your network and VPN server availability."
            error = type('TimeoutError', (Exception,), {'__str__': lambda self: error_msg})()
            self.finished_signal.emit(error)
        except Exception as e:
            self.finished_signal.emit(e)

#########################################
# Folder Browser Widget
#########################################
class FolderBrowserWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        quickLinksLayout = QHBoxLayout()
        self.homeButton = QPushButton("Home")
        self.docsButton = QPushButton("Documents")
        self.homeButton.clicked.connect(self.goHome)
        self.docsButton.clicked.connect(self.goDocuments)
        quickLinksLayout.addWidget(self.homeButton)
        quickLinksLayout.addWidget(self.docsButton)
        layout.addLayout(quickLinksLayout)
        
        self.model = QFileSystemModel()
        self.model.setRootPath(QDir.homePath())
        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(QDir.homePath()))
        self.tree.setColumnWidth(0, 200)
        layout.addWidget(self.tree)
        
        self.setLayout(layout)
        
    def goHome(self):
        self.tree.setRootIndex(self.model.index(QDir.homePath()))
    
    def goDocuments(self):
        documentsPath = os.path.join(QDir.homePath(), "Documents")
        if not os.path.exists(documentsPath):
            documentsPath = QDir.homePath()
        self.tree.setRootIndex(self.model.index(documentsPath))
    
    def setRoot(self, path):
        self.tree.setRootIndex(self.model.index(path))

#########################################
# Header Widget
#########################################
class HeaderWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("header")
        self.initUI()
    
    def initUI(self):
        layout = QHBoxLayout()
        self.companyLogoLabel = QLabel()
        companyPixmap = QPixmap("Broetje Logo.png")
        companyPixmap = companyPixmap.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.companyLogoLabel.setPixmap(companyPixmap)
        self.companyLogoLabel.setContentsMargins(80, 20, 20, 20)
        layout.addWidget(self.companyLogoLabel, alignment=Qt.AlignLeft)
        
        layout.addStretch()
        
        self.statusLabel = QLabel()
        self.loadStatusImage("none")
        self.statusLabel.setContentsMargins(20, 20, 100, 20)
        layout.addWidget(self.statusLabel, alignment=Qt.AlignRight)
        
        self.setLayout(layout)
        
    def loadStatusImage(self, status):
        if status == "german":
            pixmap = QPixmap("germany-flag.jpg")
        elif status == "us":
            pixmap = QPixmap("usa-flag.jpg")
        else:
            pixmap = QPixmap("Question_mark.jpg")
        pixmap = pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.statusLabel.setPixmap(pixmap)

#########################################
# Credentials Widget
#########################################
class CredentialsWidget(QWidget):
    connectionStatusChanged = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        
        # German Credentials Tab
        self.germanTab = QWidget()
        germanLayout = QFormLayout()
        self.germanServer = QLineEdit("vpn.broetje-automation.de")
        self.germanUsername = QLineEdit("banet.loc\\")
        self.germanPassword = QLineEdit()
        self.germanPassword.setEchoMode(QLineEdit.Password)
        germanLayout.addRow("Server:", self.germanServer)
        germanLayout.addRow("Username:", self.germanUsername)
        germanLayout.addRow("Password:", self.germanPassword)
        self.connectGermanButton = QPushButton("Connect German VPN")
        germanLayout.addRow(self.connectGermanButton)
        self.germanTab.setLayout(germanLayout)
        
        # American Credentials Tab
        self.americanTab = QWidget()
        americanLayout = QFormLayout()
        self.americanServer = QLineEdit("vpn.ba-us.com")
        self.americanUsername = QLineEdit("ba-us.com\\")
        self.americanPassword = QLineEdit()
        self.americanPassword.setEchoMode(QLineEdit.Password)
        americanLayout.addRow("Server:", self.americanServer)
        americanLayout.addRow("Username:", self.americanUsername)
        americanLayout.addRow("Password:", self.americanPassword)
        self.connectUSButton = QPushButton("Connect US VPN")
        americanLayout.addRow(self.connectUSButton)
        self.americanTab.setLayout(americanLayout)
        
        self.tabs.addTab(self.germanTab, "German Credentials")
        self.tabs.addTab(self.americanTab, "American Credentials")
        layout.addWidget(self.tabs)
        
        self.disconnectButton = QPushButton("Disconnect VPN")
        layout.addWidget(self.disconnectButton)
        
        self.setLayout(layout)
        
        self.connectGermanButton.clicked.connect(self.connectGerman)
        self.connectUSButton.clicked.connect(self.connectUS)
        self.disconnectButton.clicked.connect(self.disconnectVPN)
        
        self.tabs.currentChanged.connect(self.onTabChanged)
        
    def onTabChanged(self, index):
        if index == 0:
            self.germanUsername.setText("banet.loc\\")
        elif index == 1:
            self.americanUsername.setText("ba-us.com\\")
    
    def connectGerman(self):
        self.connectionStatusChanged.emit("german")
    
    def connectUS(self):
        self.connectionStatusChanged.emit("us")
        
    def disconnectVPN(self):
        self.connectionStatusChanged.emit("none")

#########################################
# Main Window with Progress Bar, Task Kill/Delay, and Improved Timeout/Error Handling
#########################################
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.darkMode = True
        self.setWindowTitle("VPN Manager - Dark Mode")
        self.initUI()
        self.progressTimer = QTimer()
        self.progressTimer.timeout.connect(self.updateProgress)
        self.progressValue = 0
        
    def initUI(self):
        centralWidget = QWidget()
        mainLayout = QVBoxLayout()
        
        self.header = HeaderWidget()
        mainLayout.addWidget(self.header)
        
        toggleLayout = QHBoxLayout()
        toggleLayout.addStretch()
        self.modeToggle = QCheckBox("Dark Mode")
        self.modeToggle.setChecked(True)
        self.modeToggle.stateChanged.connect(self.toggleMode)
        toggleLayout.addWidget(self.modeToggle)
        mainLayout.addLayout(toggleLayout)
        
        splitter = QSplitter(Qt.Horizontal)
        self.folderBrowser = FolderBrowserWidget()
        self.credentialsWidget = CredentialsWidget()
        splitter.addWidget(self.folderBrowser)
        splitter.addWidget(self.credentialsWidget)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        mainLayout.addWidget(splitter)
        
        self.progressBar = QProgressBar()
        self.progressBar.setRange(0, 100)
        self.progressBar.hide()
        mainLayout.addWidget(self.progressBar)
        
        self.outputBox = QTextEdit()
        self.outputBox.setObjectName("outputBox")
        self.outputBox.setReadOnly(True)
        mainLayout.addWidget(self.outputBox)
        
        centralWidget.setLayout(mainLayout)
        self.setCentralWidget(centralWidget)
        
        self.credentialsWidget.connectionStatusChanged.connect(self.attemptConnection)
        self.applyStyles()
        
    def killExistingInstance(self):
        try:
            subprocess.run(["taskkill", "/IM", "wgsslvpnc.exe", "/F"],
                           capture_output=True, text=True, timeout=10)
            self.outputBox.append("Existing VPN client instance terminated.")
        except Exception as e:
            self.outputBox.append(f"Error killing existing instance: {str(e)}")
    
    def attemptConnection(self, status):
        vpn_client_path = r"C:\Program Files (x86)\WatchGuard\WatchGuard Mobile VPN with SSL\wgsslvpnc.exe"
        
        if status == "none":
            self.outputBox.append("Attempting to disconnect VPN...")
            command = [vpn_client_path, "/disconnect"]
            try:
                result = subprocess.run(command, capture_output=True, text=True, timeout=15)
                if result.returncode == 0:
                    self.outputBox.append("VPN Disconnected.")
                    self.header.loadStatusImage("none")
                    self.folderBrowser.setRoot(QDir.homePath())
                else:
                    self.outputBox.append("Error disconnecting VPN:")
                    self.outputBox.append(result.stderr)
            except Exception as e:
                self.outputBox.append(f"Error disconnecting: {str(e)}")
            return
        
        # For connecting: kill any existing instance, then wait before starting connection.
        self.killExistingInstance()
        self.outputBox.append("Waiting 3 seconds for shutdown...")
        self.progressBar.setValue(0)
        self.progressValue = 0
        self.progressBar.show()
        self.progressTimer.start(100)
        
        QTimer.singleShot(3000, lambda: self.startConnectionWorker(status))
    
    def startConnectionWorker(self, status):
        vpn_client_path = r"C:\Program Files (x86)\WatchGuard\WatchGuard Mobile VPN with SSL\wgsslvpnc.exe"
        if status == "german":
            server = self.credentialsWidget.germanServer.text()
            username = self.credentialsWidget.germanUsername.text()
            password = self.credentialsWidget.germanPassword.text()
        elif status == "us":
            server = self.credentialsWidget.americanServer.text()
            username = self.credentialsWidget.americanUsername.text()
            password = self.credentialsWidget.americanPassword.text()
        
        self.outputBox.append(f"Attempting to connect to {status.capitalize()} VPN at {server}...")
        
        if username.endswith('\\'):
            self.outputBox.append("Note: Ensure you enter your username after the domain prefix (e.g., domain\\username).")
        
        masked_command = [
            vpn_client_path,
            "/connect",
            f"/server:{server}",
            f"/username:{username}",
            "/password:********"
        ]
        self.outputBox.append(f"Command: {' '.join(masked_command)}")
        
        command = [
            vpn_client_path,
            "/connect",
            f"/server:{server}",
            f"/username:{username}",
            f"/password:{password}"
        ]
        
        # Increase timeout to 120 seconds for slower connections.
        self.worker = ConnectionWorker(command, timeout=120)
        self.worker.finished_signal.connect(lambda res: self.onConnectionFinished(status, res))
        self.worker.start()
    
    def updateProgress(self):
        self.progressValue = (self.progressValue + 5) % 105
        self.progressBar.setValue(self.progressValue)
    
    def onConnectionFinished(self, status, result):
        self.progressTimer.stop()
        self.progressBar.hide()
        if isinstance(result, Exception):
            self.outputBox.append(f"Error connecting: {str(result)}")
            self.outputBox.append("Check your network connection and VPN server availability.")
            self.header.loadStatusImage("none")
        else:
            if result.returncode == 0:
                self.outputBox.append("Connection successful.")
                self.header.loadStatusImage(status)
                if status == "german":
                    QTimer.singleShot(5000, self.mapGermanDrive)
            else:
                self.outputBox.append("Connection failed:")
                if result.stderr:
                    self.outputBox.append(result.stderr)
                self.header.loadStatusImage("none")
    
    def mapGermanDrive(self):
        net_use_command = ["net", "use", "N:", r"\\banet.loc\baw"]
        try:
            result_net = subprocess.run(net_use_command, capture_output=True, text=True, timeout=15)
            if result_net.returncode == 0:
                self.outputBox.append("Mapped network drive N: successfully.")
                self.folderBrowser.setRoot("N:/")
            else:
                self.outputBox.append("Error mapping network drive:")
                self.outputBox.append(result_net.stderr)
        except Exception as e:
            self.outputBox.append(f"Error mapping network drive: {str(e)}")
    
    def toggleMode(self, state):
        self.darkMode = state == Qt.Checked
        if self.darkMode:
            self.setWindowTitle("VPN Manager - Dark Mode")
        else:
            self.setWindowTitle("VPN Manager - Light Mode")
        self.applyStyles()
        
    def applyStyles(self):
        if self.darkMode:
            style = """
                QMainWindow { background-color: #000000; }
                QWidget { font-family: Arial; font-size: 12pt; color: #FFFFFF; }
                QPushButton {
                    background-color: #FF0000; color: #FFFFFF;
                    border: none; padding: 6px 12px; border-radius: 4px;
                }
                QPushButton:hover { background-color: #CC0000; }
                QLineEdit {
                    border: 1px solid #FFFFFF; padding: 4px;
                    color: #FFFFFF; background-color: #333333;
                }
                QTabWidget::pane { border: 1px solid #FFFFFF; }
                QTabBar::tab {
                    background: #000000; border: 1px solid #FFFFFF;
                    padding: 8px; color: #FFFFFF;
                }
                QTabBar::tab:selected {
                    background: #FF0000; color: #FFFFFF;
                }
                QTreeView {
                    background-color: #555555; color: #FF0000;
                }
                QHeaderView::section {
                    background-color: #555555; color: #FF0000;
                }
                QTextEdit#outputBox {
                    background-color: #333333; color: #FF0000;
                    border: 1px solid #FFFFFF;
                }
                #header { background-color: #444444; }
            """
        else:
            style = """
                QMainWindow { background-color: #FFFFFF; }
                QWidget { font-family: Arial; font-size: 12pt; color: #000000; }
                QPushButton {
                    background-color: #FF0000; color: #FFFFFF;
                    border: none; padding: 6px 12px; border-radius: 4px;
                }
                QPushButton:hover { background-color: #CC0000; }
                QLineEdit {
                    border: 1px solid #000000; padding: 4px;
                    color: #000000; background-color: #FFFFFF;
                }
                QTabWidget::pane { border: 1px solid #000000; }
                QTabBar::tab {
                    background: #FFFFFF; border: 1px solid #000000;
                    padding: 8px; color: #000000;
                }
                QTabBar::tab:selected {
                    background: #FF0000; color: #FFFFFF;
                }
                QTreeView {
                    background-color: #EEEEEE; color: #000000;
                }
                QHeaderView::section {
                    background-color: #EEEEEE; color: #000000;
                }
                QTextEdit#outputBox {
                    background-color: #FFFFFF; color: #000000;
                    border: 1px solid #000000;
                }
                #header { background-color: #F0F0F0; }
            """
        self.setStyleSheet(style)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 600)
    window.show()
    sys.exit(app.exec_())
