import sys
import os
import subprocess
import shutil
import json
import re
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QHBoxLayout, QVBoxLayout, QFileSystemModel, QTreeView, QTabWidget,
    QFormLayout, QSplitter, QCheckBox, QTextEdit, QProgressBar, QMessageBox, QComboBox,
    QDialog, QListWidget, QListWidgetItem, QDialogButtonBox, QInputDialog
)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QDir, pyqtSignal, QTimer, QThread

# ------------------------------
# Dialog for adding a new server (existing, used in CustomServerDialog)
# ------------------------------
class AddServerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Server")
        self.resize(1000, 150)  # Doubled width compared to a standard dialog
        self.applyModeStyle(parent)
        self.initUI()

    def applyModeStyle(self, parent):
        if parent is not None and hasattr(parent, "darkMode") and parent.darkMode:
            style = """
                QDialog { background-color: #333333; color: #FFFFFF; }
                QLabel { color: #FFFFFF; }
                QLineEdit { background-color: #333333; color: #FFFFFF; border: 1px solid #FFFFFF; }
                QDialogButtonBox { background-color: #333333; color: #FFFFFF; }
                QMenu { background-color: #333333; color: #000000; }
            """
        else:
            style = """
                QDialog { background-color: #FFFFFF; color: #000000; }
                QLabel { color: #000000; }
                QLineEdit { background-color: #FFFFFF; color: #000000; border: 1px solid #000000; }
                QDialogButtonBox { background-color: #FFFFFF; color: #000000; }
                QMenu { background-color: #FFFFFF; color: #000000; }
            """
        self.setStyleSheet(style)

    def initUI(self):
        layout = QVBoxLayout()
        formLayout = QFormLayout()
        self.descriptionLineEdit = QLineEdit()
        self.addressLineEdit = QLineEdit()
        formLayout.addRow("Description:", self.descriptionLineEdit)
        formLayout.addRow("Server UNC Path:", self.addressLineEdit)
        layout.addLayout(formLayout)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(buttonBox)
        self.setLayout(layout)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)

    def getValues(self):
        return self.descriptionLineEdit.text().strip(), self.addressLineEdit.text().strip()

# ------------------------------
# Dialog for adding a new network folder mapping
# ------------------------------
class AddNetworkFolderDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Custom Network Folder")
        self.resize(1000, 150)
        self.applyModeStyle(parent)
        self.initUI()

    def applyModeStyle(self, parent):
        if parent is not None and hasattr(parent, "darkMode") and parent.darkMode:
            style = """
                QDialog { background-color: #333333; color: #FFFFFF; }
                QLabel { color: #FFFFFF; }
                QLineEdit { background-color: #333333; color: #FFFFFF; border: 1px solid #FFFFFF; }
                QDialogButtonBox { background-color: #333333; color: #FFFFFF; }
            """
        else:
            style = """
                QDialog { background-color: #FFFFFF; color: #000000; }
                QLabel { color: #000000; }
                QLineEdit { background-color: #FFFFFF; color: #000000; border: 1px solid #000000; }
                QDialogButtonBox { background-color: #FFFFFF; color: #000000; }
            """
        self.setStyleSheet(style)

    def initUI(self):
        layout = QVBoxLayout()
        formLayout = QFormLayout()
        self.driveLineEdit = QLineEdit()
        self.driveLineEdit.setPlaceholderText("e.g., N:")
        self.pathLineEdit = QLineEdit()
        self.pathLineEdit.setPlaceholderText("e.g., \\\\banet.loc\\baw")
        formLayout.addRow("Drive Letter:", self.driveLineEdit)
        formLayout.addRow("UNC Path:", self.pathLineEdit)
        layout.addLayout(formLayout)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(buttonBox)
        self.setLayout(layout)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)

    def getValues(self):
        drive = self.driveLineEdit.text().strip()
        path = self.pathLineEdit.text().strip()
        return drive, path

# ------------------------------
# Custom dialog for custom server locations (existing)
# ------------------------------
class CustomServerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Server Locations")
        self.resize(1000, 300)
        if parent is not None:
            self.setStyleSheet(parent.styleSheet())
        self.servers = self.loadCustomServers()
        if not self.servers:
            self.servers.append({
                "description": "PW Carrier Projects",
                "address": r"\\banet.loc\uschi\BA_Chicago\USABH_old_do_not_change\70 Projects\07_Pratt&Whitney"
            })
            self.saveCustomServers(self.servers)
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout()
        self.listWidget = QListWidget()
        if self.parent() is not None and hasattr(self.parent(), "darkMode") and self.parent().darkMode:
            self.listWidget.setStyleSheet("QListWidget { color: #FFFFFF; background-color: #333333; } QMenu { background-color: #333333; color: #000000; }")
        else:
            self.listWidget.setStyleSheet("QListWidget { color: #000000; background-color: #FFFFFF; } QMenu { background-color: #FFFFFF; color: #000000; }")
        self.refreshList()
        layout.addWidget(self.listWidget)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        addButton = QPushButton("Add New Server")
        deleteButton = QPushButton("Delete Selected")
        editButton = QPushButton("Edit Selected")
        buttonBox.addButton(addButton, QDialogButtonBox.ActionRole)
        buttonBox.addButton(deleteButton, QDialogButtonBox.ActionRole)
        buttonBox.addButton(editButton, QDialogButtonBox.ActionRole)
        layout.addWidget(buttonBox)
        self.setLayout(layout)
        addButton.clicked.connect(self.addServer)
        deleteButton.clicked.connect(self.deleteServer)
        editButton.clicked.connect(self.editServer)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)
    
    def refreshList(self):
        self.listWidget.clear()
        for server in self.servers:
            item = QListWidgetItem(server["description"])
            item.setData(Qt.UserRole, server["address"])
            self.listWidget.addItem(item)
    
    def addServer(self):
        dialog = AddServerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            desc, addr = dialog.getValues()
            if desc and addr:
                new_server = {"description": desc, "address": addr}
                self.servers.append(new_server)
                self.saveCustomServers(self.servers)
                self.refreshList()
    
    def editServer(self):
        selectedItems = self.listWidget.selectedItems()
        if not selectedItems:
            QMessageBox.warning(self, "No Selection", "Please select a server to edit.")
            return
        index = self.listWidget.currentRow()
        current_server = self.servers[index]
        # Use AddServerDialog for editing, but prepopulate with the current values
        dialog = AddServerDialog(self)
        dialog.setWindowTitle("Edit Server")
        dialog.descriptionLineEdit.setText(current_server["description"])
        dialog.addressLineEdit.setText(current_server["address"])
        if dialog.exec_() == QDialog.Accepted:
            new_desc, new_addr = dialog.getValues()
            if new_desc and new_addr:
                self.servers[index] = {"description": new_desc, "address": new_addr}
                self.saveCustomServers(self.servers)
                self.refreshList()
    
    def deleteServer(self):
        selectedItems = self.listWidget.selectedItems()
        if not selectedItems:
            return
        for item in selectedItems:
            desc = item.text().strip()
            self.servers = [s for s in self.servers if s["description"] != desc]
        self.saveCustomServers(self.servers)
        self.refreshList()
    
    def loadCustomServers(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials_manager.json")
        if os.path.exists(path):
            with open(path, "r") as f:
                try:
                    data = json.load(f)
                    return data.get("custom_servers", [])
                except Exception:
                    return []
        return []
    
    def saveCustomServers(self, servers):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials_manager.json")
        data = {}
        if os.path.exists(path):
            with open(path, "r") as f:
                try:
                    data = json.load(f)
                except Exception:
                    data = {}
        data["custom_servers"] = servers
        with open(path, "w") as f:
            json.dump(data, f, indent=4)
    
    def getSelectedServer(self):
        index = self.listWidget.currentRow()
        if index >= 0 and index < len(self.servers):
            return self.servers[index]["address"]
        return None

# ------------------------------
# Custom QComboBox that refreshes drive list when clicked
# ------------------------------
class DriveComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText("Connected Drives")
    def showPopup(self):
        if hasattr(self.parent(), "refreshDriveList"):
            self.parent().refreshDriveList()
        super().showPopup()

# ------------------------------
# Worker thread for mapping drives (verbose output)
# ------------------------------
class MappingWorker(QThread):
    finished_signal = pyqtSignal(str)
    def __init__(self, commands):
        super().__init__()
        self.commands = commands
    def run(self):
        messages = []
        for cmd in self.commands:
            messages.append("Executing: " + " ".join(cmd))
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
                messages.append("Return code: " + str(result.returncode))
                if result.stdout.strip():
                    messages.append("Output: " + result.stdout.strip())
                if result.stderr.strip():
                    messages.append("Error: " + result.stderr.strip())
                if result.returncode != 0:
                    messages.append("Error mapping drive; aborting further commands.")
                    self.finished_signal.emit("\n".join(messages))
                    return
            except Exception as e:
                messages.append("Exception: " + str(e))
                self.finished_signal.emit("\n".join(messages))
                return
        messages.append("Mapping finished successfully.")
        self.finished_signal.emit("\n".join(messages))

# ------------------------------
# Helper functions for credentials JSON
# ------------------------------
def credentials_file_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials_manager.json")

def load_credentials():
    path = credentials_file_path()
    if os.path.exists(path):
        with open(path, "r") as f:
            try:
                return json.load(f)
            except Exception:
                return {}
    return {}

def save_credentials(data):
    path = credentials_file_path()
    with open(path, "w") as f:
        json.dump(data, f, indent=4)

# ------------------------------
# Folder Browser Widget with Drive List Dropdown (and custom servers merged)
# ------------------------------
class FolderBrowserWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        quickLinksLayout = QHBoxLayout()
        self.customServersButton = QPushButton("Custom Servers")
        self.customServersButton.clicked.connect(self.showCustomServersDialog)
        openLocationButton = QPushButton("Open File Location")
        openLocationButton.clicked.connect(self.openFileLocation)
        quickLinksLayout.addWidget(self.customServersButton)
        quickLinksLayout.addWidget(openLocationButton)
        layout.addLayout(quickLinksLayout)
        self.driveComboBox = DriveComboBox(self)
        self.driveComboBox.currentIndexChanged.connect(self.driveSelected)
        layout.addWidget(self.driveComboBox)
        self.model = QFileSystemModel()
        self.model.setReadOnly(False)
        self.model.setRootPath(QDir.homePath())
        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(QDir.homePath()))
        self.tree.setColumnWidth(0, 200)
        self.tree.setEditTriggers(QTreeView.DoubleClicked | QTreeView.EditKeyPressed)
        layout.addWidget(self.tree)
        self.setLayout(layout)
        self.refreshDriveList()
        
    def openFileLocation(self):
        index = self.tree.rootIndex()
        current_path = self.model.filePath(index)
        norm_path = os.path.normpath(current_path)
        if os.path.exists(norm_path):
            os.startfile(norm_path)
        else:
            QMessageBox.warning(self, "Error", f"The current location does not exist or is not accessible:\n{norm_path}")
    
    def driveSelected(self, index):
        if index < 0:
            return
        path = self.driveComboBox.itemData(index)
        if path:
            self.setRoot(path)
        else:
            text = self.driveComboBox.currentText()
            drive = text.split()[0]
            self.setRoot(drive + "/")
    
    def setRoot(self, path):
        self.tree.setRootIndex(self.model.index(path))
        
    def refreshDriveList(self):
        items = []
        drives_added = set()
        try:
            output = subprocess.check_output(["net", "use"], text=True)
            pattern = re.compile(r"^\s*(OK|Disconnected)\s+(\w:)\s+(\\\\\S+)", re.MULTILINE)
            matches = pattern.findall(output)
            for status, drive, remote in matches:
                display = f"{drive}  {remote}"
                items.append((display, drive + "/"))
                drives_added.add(drive.upper())
        except Exception as e:
            items.append(("Error retrieving drives", ""))
        try:
            wmic_output = subprocess.check_output(["wmic", "logicaldisk", "where", "drivetype=4", "get", "DeviceID,ProviderName"], text=True)
            for line in wmic_output.splitlines()[1:]:
                if line.strip():
                    parts = line.split()
                    if len(parts) >= 2:
                        drive = parts[0]
                        provider = " ".join(parts[1:])
                        if drive.upper() not in drives_added:
                            display = f"{drive}  {provider}"
                            items.append((display, drive + "/"))
                            drives_added.add(drive.upper())
        except Exception as e:
            pass
        data = load_credentials()
        custom_servers = data.get("custom_servers", [])
        for server in custom_servers:
            desc = server.get("description", "Custom Server")
            addr = server.get("address", "")
            items.append(("Custom: " + desc, addr))
        items.sort(key=lambda x: x[0].lower())
        self.driveComboBox.clear()
        for display, item_data in items:
            self.driveComboBox.addItem(display, item_data)
    
    def showCustomServersDialog(self):
        dialog = CustomServerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.refreshDriveList()
            QMessageBox.information(self, "Custom Server", "Custom server added. Please select it from the drop down.")

# ------------------------------
# Header Widget
# ------------------------------
class HeaderWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("header")
        self.initUI()
    
    def initUI(self):
        layout = QHBoxLayout()
        self.companyLogoLabel = QLabel()
        companyPixmap = QPixmap("Broetje Logo.png")
        companyPixmap = companyPixmap.scaled(500, 500, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.companyLogoLabel.setPixmap(companyPixmap)
        self.companyLogoLabel.setContentsMargins(80, 20, 20, 20)
        layout.addWidget(self.companyLogoLabel, alignment=Qt.AlignLeft)
        layout.addStretch()
        self.statusLabel = QLabel()
        self.loadStatusImage("us")
        self.statusLabel.setContentsMargins(20, 20, 150, 20)
        layout.addWidget(self.statusLabel, alignment=Qt.AlignRight)
        self.setLayout(layout)
        
    def loadStatusImage(self, status):
        if status == "german":
            pixmap = QPixmap("germany-flag.jpg")
        elif status == "us":
            pixmap = QPixmap("usa-flag.jpg")
        else:
            pixmap = QPixmap("Question_mark.jpg")
        pixmap = pixmap.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.statusLabel.setPixmap(pixmap)

# ------------------------------
# Credentials Widget with JSON save/clear, autofill, and new network folder buttons
# ------------------------------
class CredentialsWidget(QWidget):
    connectServersRequested = pyqtSignal(str)
    rdpLaunchRequested = pyqtSignal(str)
    serverSelectionChanged = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        self.tabs.setCurrentIndex(1)
        self.loadSavedCredentials()
        self.autofillCredentials()
        self.serverSelectionChanged.emit("us")
        
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
        germanButtonsLayout = QHBoxLayout()
        self.connectGermanServersButton = QPushButton("Connect to German Network Folders")
        self.rdpGermanButton = QPushButton("German RDP Launch")
        germanButtonsLayout.addWidget(self.connectGermanServersButton)
        germanButtonsLayout.addWidget(self.rdpGermanButton)
        germanLayout.addRow(germanButtonsLayout)
        customFolderLayout = QHBoxLayout()
        self.addGermanFolderButton = QPushButton("Add Custom Network Folder")
        customFolderLayout.addWidget(self.addGermanFolderButton)
        germanLayout.addRow(customFolderLayout)
        germanCredButtonsLayout = QHBoxLayout()
        self.saveGermanButton = QPushButton("Save Credentials")
        self.clearGermanButton = QPushButton("Clear Credentials")
        germanCredButtonsLayout.addWidget(self.saveGermanButton)
        germanCredButtonsLayout.addWidget(self.clearGermanButton)
        germanLayout.addRow(germanCredButtonsLayout)
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
        americanButtonsLayout = QHBoxLayout()
        self.connectUSServersButton = QPushButton("Connect to US Network Folders")
        self.rdpUSButton = QPushButton("US RDP Launch")
        americanButtonsLayout.addWidget(self.connectUSServersButton)
        americanButtonsLayout.addWidget(self.rdpUSButton)
        americanLayout.addRow(americanButtonsLayout)
        customFolderLayout2 = QHBoxLayout()
        self.addUSFolderButton = QPushButton("Add Custom Network Folder")
        customFolderLayout2.addWidget(self.addUSFolderButton)
        americanLayout.addRow(customFolderLayout2)
        americanCredButtonsLayout = QHBoxLayout()
        self.saveAmericanButton = QPushButton("Save Credentials")
        self.clearAmericanButton = QPushButton("Clear Credentials")
        americanCredButtonsLayout.addWidget(self.saveAmericanButton)
        americanCredButtonsLayout.addWidget(self.clearAmericanButton)
        americanLayout.addRow(americanCredButtonsLayout)
        self.americanTab.setLayout(americanLayout)
        
        self.tabs.addTab(self.germanTab, "German Credentials")
        self.tabs.addTab(self.americanTab, "American Credentials")
        layout.addWidget(self.tabs)
        self.setLayout(layout)
        
        self.connectGermanServersButton.clicked.connect(lambda: self.connectServersRequested.emit("german"))
        self.rdpGermanButton.clicked.connect(lambda: self.rdpLaunchRequested.emit("german"))
        self.connectUSServersButton.clicked.connect(lambda: self.connectServersRequested.emit("us"))
        self.rdpUSButton.clicked.connect(lambda: self.rdpLaunchRequested.emit("us"))
        self.saveGermanButton.clicked.connect(self.saveGermanCredentials)
        self.clearGermanButton.clicked.connect(self.clearGermanCredentials)
        self.saveAmericanButton.clicked.connect(self.saveAmericanCredentials)
        self.clearAmericanButton.clicked.connect(self.clearAmericanCredentials)
        self.addGermanFolderButton.clicked.connect(self.addGermanNetworkFolder)
        self.addUSFolderButton.clicked.connect(self.addUSNetworkFolder)
        self.tabs.currentChanged.connect(self.onTabChanged)
        
    def onTabChanged(self, index):
        self.autofillCredentials()
        if index == 0:
            self.serverSelectionChanged.emit("german")
        elif index == 1:
            self.serverSelectionChanged.emit("us")
    
    def autofillCredentials(self):
        data = load_credentials()
        if self.tabs.currentIndex() == 0:
            if "german" in data:
                self.germanServer.setText(data["german"].get("server", "vpn.broetje-automation.de"))
                self.germanUsername.setText(data["german"].get("username", "banet.loc\\"))
                self.germanPassword.setText(data["german"].get("password", ""))
            else:
                self.germanUsername.setText("banet.loc\\")
        else:
            if "american" in data:
                self.americanServer.setText(data["american"].get("server", "vpn.ba-us.com"))
                self.americanUsername.setText(data["american"].get("username", "ba-us.com\\"))
                self.americanPassword.setText(data["american"].get("password", ""))
            else:
                self.americanUsername.setText("ba-us.com\\")
    
    def addGermanNetworkFolder(self):
        dialog = AddNetworkFolderDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            drive, path = dialog.getValues()
            if drive and path:
                data = load_credentials()
                folders = data.get("german_network_folders", [])
                folders.append({"drive": drive, "path": path})
                data["german_network_folders"] = folders
                save_credentials(data)
                QMessageBox.information(self, "Network Folder", "German network folder added.")
    
    def addUSNetworkFolder(self):
        dialog = AddNetworkFolderDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            drive, path = dialog.getValues()
            if drive and path:
                data = load_credentials()
                folders = data.get("american_network_folders", [])
                folders.append({"drive": drive, "path": path})
                data["american_network_folders"] = folders
                save_credentials(data)
                QMessageBox.information(self, "Network Folder", "US network folder added.")
    
    def saveGermanCredentials(self):
        data = load_credentials()
        data["german"] = {
            "server": self.germanServer.text(),
            "username": self.germanUsername.text(),
            "password": self.germanPassword.text()
        }
        save_credentials(data)
        QMessageBox.information(self, "Save Credentials", "German credentials saved.")
    
    def clearGermanCredentials(self):
        self.germanServer.clear()
        self.germanUsername.clear()
        self.germanPassword.clear()
        QMessageBox.information(self, "Clear Credentials", "German credentials cleared.")
    
    def saveAmericanCredentials(self):
        data = load_credentials()
        data["american"] = {
            "server": self.americanServer.text(),
            "username": self.americanUsername.text(),
            "password": self.americanPassword.text()
        }
        save_credentials(data)
        QMessageBox.information(self, "Save Credentials", "American credentials saved.")
    
    def clearAmericanCredentials(self):
        self.americanServer.clear()
        self.americanUsername.clear()
        self.americanPassword.clear()
        QMessageBox.information(self, "Clear Credentials", "American credentials cleared.")
    
    def loadSavedCredentials(self):
        data = load_credentials()
        if "german" in data:
            self.germanServer.setText(data["german"].get("server", "vpn.broetje-automation.de"))
            self.germanUsername.setText(data["german"].get("username", "banet.loc\\"))
            self.germanPassword.setText(data["german"].get("password", ""))
        if "american" in data:
            self.americanServer.setText(data["american"].get("server", "vpn.ba-us.com"))
            self.americanUsername.setText(data["american"].get("username", "ba-us.com\\"))
            self.americanPassword.setText(data["american"].get("password", ""))

# ------------------------------
# Main Window with File Browser, Network Folder Connect, RDP Launch, VPN Controls, and Custom Servers
# ------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.darkMode = True
        self.setWindowTitle("VPN Manager - Dark Mode")
        self.currentMappingStatus = None
        self.initUI()
        
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
        vpnButtonsLayout = QHBoxLayout()
        self.openWatchGuardButton = QPushButton("Open WatchGuard VPN")
        self.disconnectButton = QPushButton("Disconnect VPN")
        vpnButtonsLayout.addWidget(self.openWatchGuardButton)
        vpnButtonsLayout.addWidget(self.disconnectButton)
        vpnButtonsLayout.setStretch(0, 1)
        vpnButtonsLayout.setStretch(1, 1)
        mainLayout.addLayout(vpnButtonsLayout)
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
        self.credentialsWidget.serverSelectionChanged.connect(self.header.loadStatusImage)
        self.credentialsWidget.connectServersRequested.connect(self.connectNetworkFolders)
        self.credentialsWidget.rdpLaunchRequested.connect(self.launchRDP)
        self.openWatchGuardButton.clicked.connect(self.openWatchGuard)
        self.disconnectButton.clicked.connect(self.disconnectVPN)
        self.applyStyles()
        
    def openWatchGuard(self):
        vpn_client_path = r"C:\Program Files (x86)\WatchGuard\WatchGuard Mobile VPN with SSL\wgsslvpnc.exe"
        try:
            tasklist = subprocess.check_output(["tasklist", "/FI", "IMAGENAME eq wgsslvpnc.exe"], text=True)
            if "wgsslvpnc.exe" in tasklist:
                QMessageBox.information(self, "Process Running", "WatchGuard is already running.")
                return
        except Exception as e:
            self.outputBox.append(f"Error checking process: {str(e)}")
        self.outputBox.append("Launching WatchGuard application...")
        try:
            subprocess.Popen(vpn_client_path)
            self.outputBox.append("WatchGuard application launched successfully.")
        except Exception as e:
            self.outputBox.append(f"Error launching WatchGuard: {str(e)}")
    
    def disconnectVPN(self):
        self.outputBox.append("Attempting to disconnect VPN and terminate WatchGuard...")
        try:
            subprocess.run(["taskkill", "/IM", "wgsslvpnc.exe", "/F"],
                           capture_output=True, text=True, timeout=10)
            self.outputBox.append("WatchGuard application terminated.")
            self.folderBrowser.setRoot(QDir.homePath())
        except Exception as e:
            self.outputBox.append(f"Error disconnecting: {str(e)}")
    
    def connectNetworkFolders(self, status):
        self.currentMappingStatus = status
        data = load_credentials()
        disconnectCommands = []
        connectionCommands = []
        if status == "german":
            self.credentialsWidget.connectGermanServersButton.setEnabled(False)
            folders = data.get("german_network_folders", [])
            if not folders:
                folders = [
                    {"drive": "N:", "path": r"\\banet.loc\baw"},
                    {"drive": "I:", "path": r"\\banet.loc\derae.user\home"}
                ]
            username = self.credentialsWidget.germanUsername.text()
            password = self.credentialsWidget.germanPassword.text()
            for mapping in folders:
                drive = mapping.get("drive")
                path = mapping.get("path")
                connectionCommands.append(["net", "use", drive, path, f"/user:{username}", password])
                disconnectCommands.append(["net", "use", drive, "/delete", "/Y"])
        elif status == "us":
            self.credentialsWidget.connectUSServersButton.setEnabled(False)
            folders = data.get("american_network_folders", [])
            if not folders:
                folders = [
                    {"drive": "Z:", "path": r"\\fs02\uschi"}
                ]
            full_username = self.credentialsWidget.americanUsername.text()
            password = self.credentialsWidget.americanPassword.text()
            username_extracted = full_username.split('\\')[-1]
            for mapping in folders:
                drive = mapping.get("drive")
                path = mapping.get("path")
                final_username = f"BA-US\\{username_extracted}"
                connectionCommands.append(["net", "use", drive, path, f"/user:{final_username}", password])
                disconnectCommands.append(["net", "use", drive, "/delete", "/Y"])
        else:
            self.outputBox.append("Unknown network folder selection for connection.")
            return

        self.outputBox.append("Starting disconnection of existing network folders...")
        # Define a recursive function to disconnect drives one by one.
        def disconnectNext(cmds, idx, onComplete):
            if idx < len(cmds):
                cmd = cmds[idx]
                drive = cmd[2] if len(cmd) >= 3 else "Unknown"
                self.outputBox.append(f"Disconnecting drive {cmd[2].replace('/', '')}...")
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
                    self.outputBox.append(f"Disconnected drive {cmd[2].replace('/', '')}. Return code: {result.returncode}")
                except Exception as e:
                    self.outputBox.append(f"Exception disconnecting drive {cmd[2].replace('/', '')}: {str(e)}")
                # Wait 2 seconds before disconnecting the next drive.
                QTimer.singleShot(2000, lambda: disconnectNext(cmds, idx+1, onComplete))
            else:
                onComplete()

        # After disconnecting, wait 3 seconds and then start connecting.
        def startConnecting():
            self.outputBox.append("All network folders disconnected. Starting reconnection...")
            self.mappingWorker = MappingWorker(connectionCommands)
            self.mappingWorker.finished_signal.connect(self.mappingFinished)
            self.mappingWorker.start()
        
        disconnectNext(disconnectCommands, 0, lambda: QTimer.singleShot(3000, startConnecting))
    
    def mappingFinished(self, msg):
        self.outputBox.append(msg)
        QTimer.singleShot(2000, self.folderBrowser.refreshDriveList)
        if self.currentMappingStatus == "german":
            self.credentialsWidget.connectGermanServersButton.setEnabled(True)
        elif self.currentMappingStatus == "us":
            self.credentialsWidget.connectUSServersButton.setEnabled(True)
    
    def launchRDP(self, status):
        if status == "german":
            url = "https://rds.banet.loc/RDWeb/Pages/en-US/Default.aspx"
        elif status == "us":
            url = "https://rds.ba-us.com/RDWeb/Pages/en-US/Default.aspx"
        else:
            self.outputBox.append("Unknown network folder selection for RDP launch.")
            return
        chrome_path = shutil.which("chrome.exe")
        if not chrome_path:
            possible_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    chrome_path = path
                    break
        if chrome_path:
            self.outputBox.append(f"Launching Chrome to {url}...")
            try:
                subprocess.Popen([chrome_path, url])
                self.outputBox.append("Chrome launched successfully.")
            except Exception as e:
                self.outputBox.append(f"Error launching Chrome: {str(e)}")
        else:
            QMessageBox.warning(self, "Chrome Not Found", "Google Chrome is not installed. Please install Chrome.")
    
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
                QComboBox {
                    background-color: #333333;
                    color: #FF0000;
                    border: 1px solid #FFFFFF;
                    padding: 2px;
                }
                QComboBox QAbstractItemView {
                    background-color: #333333;
                    color: #FF0000;
                    selection-background-color: #FF0000;
                    selection-color: #FFFFFF;
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
                QComboBox {
                    background-color: #FFFFFF;
                    color: #000000;
                    border: 1px solid #000000;
                    padding: 2px;
                }
                QComboBox QAbstractItemView {
                    background-color: #FFFFFF;
                    color: #000000;
                    selection-background-color: #FF0000;
                    selection-color: #FFFFFF;
                }
                #header { background-color: #F0F0F0; }
            """
        self.setStyleSheet(style)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1500, 1000)
    window.show()
    sys.exit(app.exec_())
