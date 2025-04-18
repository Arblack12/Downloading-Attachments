# main.py

import sys
import os
import logging
import configparser
import re
import traceback
from datetime import datetime, date
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QSystemTrayIcon
from imap_tools import MailBox
from logging.handlers import RotatingFileHandler
import importlib
import json
import shutil

################################################################################
#                          VERSION & CHANGELOG SECTION                         #
################################################################################

VERSION = '0.6.0'
CHANGELOG = r"""
# Worms Direct Management Changelog

## Version 0.6.0

- Integrated invoice-processing functionality fully into the main script.
- Removed references to a separate "smaller script" for clarity.
- Unified datetime usage to avoid conflicts (no more "module 'datetime' has no attribute 'now'" errors).

## Version 0.5.0

- Transitioned configuration management from `config.ini` to the GUI.
- Implemented dynamic loading of scripts as tabs in the main window.
- Improved error handling and logging.

## [0.4.0] - 2024-10-09
### Changed
- Updated polling interval from 5 minutes (300 seconds) to 10 minutes (600 seconds) to reduce the frequency of IMAP checks and minimize the risk of being flagged by Apple.
- Changed backup folder location from `D:\Sync\Businesses\Worms Direct\Worms Direct Software\Invoices\Temporary backups` to `D:\Sync\Businesses\Worms Direct\Invoices\Temporary backups`.

### Added
- Implemented a main GUI window that can be toggled via double-clicking the system tray icon.
- Enhanced system tray interactions with double-click functionality and additional menu options.
- Integrated version tracking in the system tray tooltip.
- Modularized the application by separating functionalities into different scripts loaded as tabs.
- Added **Task Automator** tab for automated backup tasks, including file backups and deletion of old backups.

### Fixed
- Resolved intermittent SSL errors by implementing exponential backoff more effectively.
- Ensured that the main window hides instead of closing, preventing the application from exiting unintentionally.

## [0.3.0] - 2024-10-08
### Added
- Implemented exponential backoff mechanism to handle connection failures gracefully.
- Enhanced logging with rotating file handlers to manage log file sizes.
- Added system tray GUI with context menu options:
  - View Logs
  - View Changelog
  - Check Now
  - Pause Monitoring
  - Resume Monitoring
  - Open Download Folder
  - Exit

### Fixed
- Resolved issue where duplicate `password` entries in `config.ini` caused `DuplicateOptionError`.
- Ensured that attachments are saved directly into the specified download folder without creating subdirectories.

## [0.2.0] - 2024-10-08
### Added
- Configured the script to run silently using `pythonw.exe`, preventing the CMD window from appearing.
- Set the polling interval to 5 minutes (300 seconds) to balance server load and responsiveness.
- Created a separate `CHANGELOG.md` file to track application updates and versions.

### Fixed
- Addressed the error caused by incorrect usage of the `AND` operator in the `imap_tools` library.
- Ensured that the download folder path in `config.ini` is correctly specified and utilized by the script.

## [0.1.0] - 2024-10-08
### Added
- Initial release of the application to download attachments from the "Invoices" IMAP folder.
- Configured `config.ini` with essential settings:
  - Email account credentials
  - IMAP server details
  - Download folder path
  - Polling interval
  - Logging settings
- Implemented basic error handling and logging mechanisms.
- Set up the application to run as a system tray application using PyQt5.
"""

################################################################################
#                                 MAIN WINDOW                                  #
################################################################################

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, tray_icon, config, parent=None):
        super(MainWindow, self).__init__(parent)
        self.tray_icon = tray_icon
        self.config = config
        self.setWindowTitle(f"Worms Direct Management v{VERSION}")
        self.setGeometry(400, 400, 800, 600)
        
        # Central Widget with Tabs
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout
        layout = QtWidgets.QVBoxLayout()
        
        # Tab Widget
        self.tabs = QtWidgets.QTabWidget()
        
        # Dynamically load scripts as tabs
        self.load_scripts()
        
        # Add the "Invoices Management" tab
        from __main__ import InvoicesManagementTab
        self.tabs.addTab(InvoicesManagementTab(self.config), "Invoices Management")
        
        # Add tabs to the main layout
        layout.addWidget(self.tabs)
        central_widget.setLayout(layout)
    
    def load_scripts(self):
        """Dynamically load scripts from the 'scripts' directory as tabs."""
        scripts_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'scripts')
        if not os.path.exists(scripts_path):
            logging.error(f"Scripts directory not found at {scripts_path}")
            return
        
        sys.path.insert(0, scripts_path)  # Add scripts directory to path
        
        for filename in os.listdir(scripts_path):
            if filename.endswith('.py') and filename != '__init__.py':
                module_name = filename[:-3]
                try:
                    module = importlib.import_module(module_name)
                    if hasattr(module, 'Tab'):
                        tab_class = getattr(module, 'Tab')
                        tab_instance = tab_class(self.config)
                        self.tabs.addTab(tab_instance, tab_instance.tab_name)
                        logging.info(f"Loaded script '{module_name}' as tab '{tab_instance.tab_name}'.")
                    else:
                        logging.warning(f"Module '{module_name}' does not have a 'Tab' class.")
                except Exception as e:
                    logging.error(f"Failed to load script '{module_name}': {traceback.format_exc()}")
    
    def closeEvent(self, event):
        """Override the close event to hide the window instead of closing."""
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "Worms Direct Management",
            "Application was minimized to the system tray.",
            QtWidgets.QSystemTrayIcon.Information,
            2000
        )
        logging.info("Main window hidden instead of closed.")

################################################################################
#                                  LOG WINDOW                                  #
################################################################################

class LogWindow(QtWidgets.QWidget):
    def __init__(self, log_file, parent=None):
        super(LogWindow, self).__init__(parent)
        self.setWindowTitle("Worms Direct Management Logs")
        self.setGeometry(300, 300, 600, 400)
        self.log_file = log_file
        
        # Layout
        layout = QtWidgets.QVBoxLayout()
        
        # Text Edit for logs
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        
        # Load log file
        self.refresh_logs()
        
        # Refresh Button
        self.refresh_button = QtWidgets.QPushButton("Refresh Logs")
        self.refresh_button.clicked.connect(self.refresh_logs)
        
        layout.addWidget(self.log_text)
        layout.addWidget(self.refresh_button)
        self.setLayout(layout)
    
    def refresh_logs(self):
        """Refresh the log display."""
        if os.path.exists(self.log_file):
            try:
                with open(self.log_file, 'r') as f:
                    self.log_text.setPlainText(f.read())
                logging.info("Log window refreshed.")
            except Exception as e:
                self.log_text.setPlainText(f"Failed to refresh logs: {e}")
                logging.error(f"Failed to refresh logs in log window: {e}")
        else:
            self.log_text.setPlainText("No logs available.")
            logging.warning("Log window refresh attempted but no log file found.")

################################################################################
#                                CHANGELOG WINDOW                              #
################################################################################

class ChangelogWindow(QtWidgets.QWidget):
    def __init__(self, changelog, parent=None):
        super(ChangelogWindow, self).__init__(parent)
        self.setWindowTitle("Changelog")
        self.setGeometry(350, 350, 700, 500)
        
        # Layout
        layout = QtWidgets.QVBoxLayout()
        
        # Text Edit for changelog
        self.changelog_text = QtWidgets.QTextEdit()
        self.changelog_text.setReadOnly(True)
        self.changelog_text.setPlainText(changelog)
        
        # Close Button
        self.close_button = QtWidgets.QPushButton("Close")
        self.close_button.clicked.connect(self.close)
        
        layout.addWidget(self.changelog_text)
        layout.addWidget(self.close_button)
        self.setLayout(layout)

################################################################################
#                           SYSTEM TRAY APPLICATION                            #
################################################################################

class EmailAttachmentDownloader(QtWidgets.QSystemTrayIcon):
    def __init__(self, icon, config, parent=None):
        super(EmailAttachmentDownloader, self).__init__(icon, parent)
        self.setToolTip(f"Worms Direct Management v{VERSION}")
        
        # Load configuration
        self.config = config
        self.EMAIL_ACCOUNT = self.config.get('EMAIL', 'ACCOUNT', fallback='')
        self.PASSWORD = self.get_password()
        self.IMAP_SERVER = self.config.get('IMAP', 'SERVER', fallback='')
        self.IMAP_PORT = int(self.config.get('IMAP', 'PORT', fallback='993'))
        self.FOLDER = self.config.get('IMAP', 'FOLDER', fallback='Invoices')
        self.download_folder = self.config.get('DOWNLOAD', 'FOLDER_PATH', fallback='D:\\Sync\\Businesses\\Worms Direct\\Invoices')
        os.makedirs(self.download_folder, exist_ok=True)
        
        # Exponential Backoff
        self.initial_backoff = int(self.config.get('SETTINGS', 'INITIAL_BACKOFF', fallback='60'))
        self.backoff_factor = int(self.config.get('SETTINGS', 'BACKOFF_FACTOR', fallback='2'))
        self.max_backoff = int(self.config.get('SETTINGS', 'MAX_BACKOFF', fallback='3600'))
        self.current_backoff = self.initial_backoff
        self.failure_count = 0
        
        # Polling Interval
        self.polling_interval = int(self.config.get('SETTINGS', 'POLLING_INTERVAL', fallback='600'))
        
        # Logging Setup
        self.setup_logging(self.config.get('SETTINGS', 'LOG_FILE', fallback='invoice_downloader.log'))
        logging.info(f"Worms Direct Management v{VERSION} started.")
        
        # Processed UIDs Setup
        self.processed_uids_file = self.config.get('SETTINGS', 'PROCESSED_UIDS_FILE', fallback='processed_uids.txt')
        self.processed_uids = set()
        self.load_processed_uids()
        
        # Timer Setup
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.check_emails)
        self.timer.start(self.polling_interval * 1000)
        
        # Windows
        self.log_window = None
        self.changelog_window = None
        
        # Monitoring
        self.monitoring_paused = False
        
        # Main Window Setup
        self.main_window = MainWindow(self, config)
        
        # System Tray Menu
        self.create_tray_menu()
        
        # Double-click to toggle
        self.activated.connect(self.on_tray_activated)
        
        # Immediately check for emails
        self.check_emails()
    
    def setup_logging(self, log_file):
        """Configure logging with rotating file handler."""
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        
        # Remove default handlers
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        
        handler = RotatingFileHandler(
            log_file,
            maxBytes=5 * 1024 * 1024,
            backupCount=3
        )
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        
        logger.addHandler(handler)
        logging.info("Logging initialized.")
    
    def get_password(self):
        """Retrieve password from config, possibly using env vars."""
        password = self.config.get('EMAIL', 'PASSWORD', fallback='')
        match = re.match(r'\$\{(.+)\}', password)
        if match:
            env_var = match.group(1)
            env_password = os.getenv(env_var)
            if env_password:
                return env_password
            else:
                logging.warning(f"Environment variable {env_var} not set for EMAIL.PASSWORD.")
                return ""
        return password
    
    def load_processed_uids(self):
        """Load processed email UIDs from a file."""
        if os.path.exists(self.processed_uids_file):
            try:
                with open(self.processed_uids_file, 'r') as f:
                    for line in f:
                        uid = line.strip()
                        if uid.isdigit():
                            self.processed_uids.add(uid)
                    logging.info("Loaded processed UIDs.")
            except Exception as e:
                logging.error(f"Failed to load processed UIDs: {e}")
        else:
            logging.info("No processed UIDs file found. Starting fresh.")
    
    def save_processed_uid(self, uid):
        """Save a processed email UID."""
        try:
            with open(self.processed_uids_file, 'a') as f:
                f.write(f"{uid}\n")
            self.processed_uids.add(uid)
            logging.info(f"Saved processed UID: {uid}")
        except Exception as e:
            logging.error(f"Failed to save processed UID {uid}: {e}")
    
    def sanitize_filename(self, filename):
        """Make filename safe for Windows."""
        invalid_chars = r'[<>:"/\\|?*]'
        sanitized = re.sub(invalid_chars, '_', filename).strip()
        return sanitized
    
    def create_tray_menu(self):
        """System tray context menu."""
        self.menu = QtWidgets.QMenu()
        
        # View Logs
        view_logs_action = self.menu.addAction("View Logs")
        view_logs_action.triggered.connect(self.show_log_window)
        
        # Changelog
        view_changelog_action = self.menu.addAction("View Changelog")
        view_changelog_action.triggered.connect(self.show_changelog_window)
        
        # Check Now
        check_now_action = self.menu.addAction("Check Now")
        check_now_action.triggered.connect(self.check_emails)
        
        # Pause/Resume
        pause_action = self.menu.addAction("Pause Monitoring")
        pause_action.triggered.connect(self.pause_monitoring)
        
        resume_action = self.menu.addAction("Resume Monitoring")
        resume_action.triggered.connect(self.resume_monitoring)
        
        # Open Download Folder
        open_folder_action = self.menu.addAction("Open Download Folder")
        open_folder_action.triggered.connect(self.open_download_folder)
        
        # Show Main Window
        show_main_window_action = self.menu.addAction("Show Main Window")
        show_main_window_action.triggered.connect(self.show_main_window)
        
        # Exit
        exit_action = self.menu.addAction("Exit")
        exit_action.triggered.connect(QtWidgets.qApp.quit)
        
        self.setContextMenu(self.menu)
    
    def on_tray_activated(self, reason):
        """Toggle main window on double-click."""
        if reason == QtWidgets.QSystemTrayIcon.DoubleClick:
            if self.main_window.isVisible():
                self.main_window.hide()
            else:
                self.main_window.show()
                self.main_window.raise_()
                self.main_window.activateWindow()
    
    def check_emails(self):
        """Fetch new emails, download attachments."""
        if self.monitoring_paused:
            logging.info("Monitoring is paused. Skipping email check.")
            return
        
        logging.info("Checking for new emails...")
        try:
            with MailBox(self.IMAP_SERVER, self.IMAP_PORT).login(self.EMAIL_ACCOUNT, self.PASSWORD) as mailbox:
                mailbox.folder.set(self.FOLDER)
                logging.info(f"Connected to folder: {self.FOLDER}")
                
                emails = mailbox.fetch()  # both read/unread
                new_attachments = 0
                
                for msg in emails:
                    uid = str(msg.uid)
                    if uid in self.processed_uids:
                        continue
                    
                    if not msg.attachments:
                        logging.info(f"No attachments in UID {uid}. Skipping.")
                        self.save_processed_uid(uid)
                        continue
                    
                    subject = msg.subject if msg.subject else "No_Subject"
                    logging.info(f"Processing email UID: {uid}, Subject: {subject}")
                    
                    for att in msg.attachments:
                        original_filename = self.sanitize_filename(att.filename or "")
                        if not original_filename:
                            logging.warning("Attachment with no filename. Skipping.")
                            continue
                        
                        sender_email = msg.from_
                        if isinstance(sender_email, tuple):
                            sender_email = sender_email[1]
                        sender_email = self.sanitize_filename(sender_email)
                        
                        invoice_name = self.sanitize_filename(subject)
                        
                        date_str = datetime.now().strftime("%Y%m%d")
                        unique_id = uid
                        
                        new_filename = f"{sender_email}_{invoice_name}_{date_str}_{unique_id}_{original_filename}"
                        filepath = os.path.join(self.download_folder, new_filename)
                        
                        if os.path.exists(filepath):
                            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                            name, ext = os.path.splitext(new_filename)
                            new_filename = f"{name}_{timestamp}{ext}"
                            filepath = os.path.join(self.download_folder, new_filename)
                        
                        # Save
                        try:
                            with open(filepath, "wb") as f:
                                f.write(att.payload)
                            logging.info(f"Downloaded: {new_filename}")
                            new_attachments += 1
                        except Exception as e:
                            logging.error(f"Failed saving {new_filename}: {e}")
                    
                    self.save_processed_uid(uid)
                
                if new_attachments > 0:
                    self.showMessage(
                        "Worms Direct Management",
                        f"Downloaded {new_attachments} new attachment(s).",
                        QtWidgets.QSystemTrayIcon.Information,
                        5000
                    )
                    logging.info(f"Downloaded {new_attachments} new attachment(s).")
                    # Reset backoff
                    self.failure_count = 0
                    self.current_backoff = self.initial_backoff
                else:
                    logging.info("No new attachments found.")
        
        except Exception as e:
            logging.error(f"An error occurred while checking emails: {e}")
            self.showMessage(
                "Worms Direct Management",
                f"Error: {e}",
                QtWidgets.QSystemTrayIcon.Critical,
                5000
            )
            # Exponential backoff
            self.failure_count += 1
            self.current_backoff = min(
                self.initial_backoff * (self.backoff_factor ** self.failure_count),
                self.max_backoff
            )
            logging.info(f"Applying backoff: {self.current_backoff} seconds")
            self.timer.stop()
            self.timer.start(self.current_backoff * 1000)
    
    def show_log_window(self):
        """Display logs."""
        if self.log_window is None:
            log_file = self.config.get('SETTINGS', 'LOG_FILE', fallback='invoice_downloader.log')
            self.log_window = LogWindow(log_file)
        self.log_window.show()
        self.log_window.raise_()
        self.log_window.activateWindow()
    
    def show_changelog_window(self):
        """Show changelog."""
        if self.changelog_window is None:
            self.changelog_window = ChangelogWindow(CHANGELOG)
        self.changelog_window.show()
        self.changelog_window.raise_()
        self.changelog_window.activateWindow()
        logging.info("Displayed changelog.")
    
    def pause_monitoring(self):
        """Pause email checking."""
        if not self.monitoring_paused:
            self.timer.stop()
            self.monitoring_paused = True
            self.showMessage("Worms Direct Management", "Monitoring paused.", QtWidgets.QSystemTrayIcon.Warning, 3000)
            logging.info("Monitoring paused.")
    
    def resume_monitoring(self):
        """Resume email checking."""
        if self.monitoring_paused:
            self.timer.start(self.polling_interval * 1000)
            self.monitoring_paused = False
            self.showMessage("Worms Direct Management", "Monitoring resumed.", QtWidgets.QSystemTrayIcon.Information, 3000)
            logging.info("Monitoring resumed.")
    
    def open_download_folder(self):
        """Open the invoice download folder."""
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(self.download_folder))
        logging.info("Opened download folder.")
    
    def show_main_window(self):
        """Show main window."""
        self.main_window.show()
        self.main_window.raise_()
        self.main_window.activateWindow()
    
    def showMessage(self, title, message, icon, msecs=5000):
        """Override to log tray notifications."""
        super().showMessage(title, message, icon, msecs)
        logging.info(f"Notification - {title}: {message}")

################################################################################
#                               CONFIG LOADER                                  #
################################################################################

def load_config(config_file=r'D:\Sync\Businesses\Worms Direct\Scripts\Downloading Attachments\config.ini'):
    config = configparser.ConfigParser()
    if not os.path.exists(config_file):
        print(f"Configuration file {config_file} not found.")
        logging.critical(f"Configuration file {config_file} not found.")
        sys.exit(1)
    config.read(config_file)
    
    for section in config.sections():
        for key in config[section]:
            value = config[section][key]
            if '#' in value:
                value = value.split('#', 1)[0].strip()
                config[section][key] = value
            
            match = re.match(r'\$\{(.+)\}', value)
            if match:
                var_name = match.group(1)
                env_value = os.getenv(var_name)
                if env_value:
                    config[section][key] = env_value
                else:
                    logging.warning(f"Environment variable {var_name} not set for {section}.{key}.")
    return config

################################################################################
#                              INVOICE PROCESSING                              #
################################################################################

# The base directory for invoice files:
base_dir = r'D:\Sync\Businesses\Worms Direct\Invoices'

# Pattern to extract email from the filename:
email_pattern = r'^([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)_'

# We'll dynamically load rename rules from invoices_config.json
companies = {}

def load_companies_from_json():
    """
    Read invoices_config.json (list of dicts) and build a dictionary.
    Each entry:
      {
        "sender_email": "some@domain.com",
        "folder_name": "Something",
        "file_name": "FilePrefix",
        "month_offset": 0,
        "day_offset": 0
      }
    We'll produce a dict keyed by sender_email -> rename rules.
    """
    script_dir = os.path.dirname(__file__)
    config_path = os.path.join(script_dir, 'invoices_config.json')
    if not os.path.exists(config_path):
        print("No invoices_config.json found. Using empty dictionary.")
        return {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Failed to load invoices_config.json: {e}")
        return {}
    
    result = {}
    for item in data:
        sender = item.get('sender_email', '').lower()
        folder_name = item.get('folder_name') or ''
        file_name = item.get('file_name') or 'NoName'
        month_offset = item.get('month_offset', 0)
        day_offset = item.get('day_offset', 0)
        
        use_sub = bool(folder_name)
        
        result[sender] = {
            'name': file_name,
            'folder_name': folder_name,
            'month_offset': month_offset,
            'day_offset': day_offset,
            'use_subfolder': use_sub,
            'numbered_files': True
        }
    return result

def get_target_date(month_offset, day_offset=0):
    """
    Start from today's date, offset by month_offset and day_offset.
    If the resulting day is out of range, fallback to the 1st of that month.
    """
    today = date.today()
    year = today.year
    month = today.month + month_offset
    day_ = today.day + day_offset
    
    while month > 12:
        month -= 12
        year += 1
    while month < 1:
        month += 12
        year -= 1
    
    try:
        return date(year, month, day_)
    except ValueError:
        return date(year, month, 1)

def process_files():
    """
    Reload rename rules from JSON, then rename and move any matching files in base_dir.
    """
    global companies
    companies = load_companies_from_json()
    
    for filename in os.listdir(base_dir):
        file_path = os.path.join(base_dir, filename)
        if not os.path.isfile(file_path):
            print(f"Skipping non-file: {filename}")
            continue
        
        match = re.match(email_pattern, filename)
        if not match:
            print(f"Email address not found in filename: {filename}")
            continue
        
        email_address = match.group(1).lower()
        if email_address not in companies:
            print(f"Email address not recognized: {email_address}")
            continue
        
        rules = companies[email_address]
        company_name = rules['name']
        folder_name = rules['folder_name']
        month_offset = rules['month_offset']
        day_offset = rules['day_offset']
        use_subfolder = rules['use_subfolder']
        numbered_files = rules['numbered_files']
        
        target_date = get_target_date(month_offset, day_offset)
        month_str = target_date.strftime('%B')
        year_str = str(target_date.year)
        
        if use_subfolder and folder_name:
            target_folder = os.path.join(base_dir, year_str, month_str, folder_name)
        else:
            target_folder = os.path.join(base_dir, year_str, month_str)
        
        os.makedirs(target_folder, exist_ok=True)
        
        ext = os.path.splitext(filename)[1]
        if numbered_files:
            existing_files = os.listdir(target_folder)
            pattern = re.compile(r'^{}-(\d+)\b'.format(re.escape(company_name)))
            numbers = []
            for f in existing_files:
                m = pattern.match(f)
                if m:
                    numbers.append(int(m.group(1)))
            next_number = max(numbers, default=0) + 1
            new_filename = f"{company_name}-{next_number}{ext}"
        else:
            new_filename = f"{company_name}{ext}"
        
        new_file_path = os.path.join(target_folder, new_filename)
        if os.path.exists(new_file_path):
            print(f"File exists (not overwritten): {new_file_path}")
            continue
        
        shutil.move(file_path, new_file_path)
        print(f"Moved file to: {new_file_path}")

################################################################################
#                   INVOICES MANAGEMENT TAB (GUI FOR JSON FILES)               #
################################################################################

class InvoicesManagementTab(QtWidgets.QWidget):
    """
    A tab for editing invoices_config.json, ignoring patterns, etc.
    Has a "Process Invoices" button that calls process_files().
    """
    def __init__(self, config, parent=None):
        super(InvoicesManagementTab, self).__init__(parent)
        self.config = config
        self.tab_name = "Invoices Management"

        # JSON file paths
        self.invoices_config_path = os.path.join(os.path.dirname(__file__), 'invoices_config.json')
        self.invoices_ignore_path = os.path.join(os.path.dirname(__file__), 'invoices_ignore.json')
        self.processed_hashes_path = os.path.join(os.path.dirname(__file__), 'processed_hashes.json')

        # Load or create data
        self.invoices_config_data = self.load_json(self.invoices_config_path)
        if not isinstance(self.invoices_config_data, list):
            self.invoices_config_data = []
        
        self.ignore_files_data = self.load_json(self.invoices_ignore_path)
        if not isinstance(self.ignore_files_data, list):
            self.ignore_files_data = []
        
        self.processed_hashes_data = self.load_json(self.processed_hashes_path)
        if not isinstance(self.processed_hashes_data, dict):
            self.processed_hashes_data = {}

        # Layout
        main_layout = QtWidgets.QVBoxLayout()

        # Title label
        title_label = QtWidgets.QLabel("<h1>Invoices Management</h1>")
        main_layout.addWidget(title_label)

        desc_label = QtWidgets.QLabel(
            "Manage downloaded invoice attachments.\n"
            "- Add/edit sender email rules.\n"
            "- Optionally specify subfolders.\n"
            "- 'Process Invoices' applies these rules immediately."
        )
        main_layout.addWidget(desc_label)

        # Sender Email Config group
        sender_group = QtWidgets.QGroupBox("Sender Email Configurations")
        sender_layout = QtWidgets.QVBoxLayout()

        self.sender_table = QtWidgets.QTableWidget()
        self.sender_table.setColumnCount(5)
        self.sender_table.setHorizontalHeaderLabels(
            ["Sender Email", "Folder Name", "File Name", "Month Offset", "Day Offset"]
        )
        self.sender_table.horizontalHeader().setStretchLastSection(True)
        self.sender_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.sender_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        sender_layout.addWidget(self.sender_table)

        self.populate_sender_table()

        btn_layout = QtWidgets.QHBoxLayout()
        self.add_sender_btn = QtWidgets.QPushButton("Add")
        self.edit_sender_btn = QtWidgets.QPushButton("Edit")
        self.remove_sender_btn = QtWidgets.QPushButton("Remove")
        self.add_sender_btn.clicked.connect(self.add_sender_config)
        self.edit_sender_btn.clicked.connect(self.edit_sender_config)
        self.remove_sender_btn.clicked.connect(self.remove_sender_config)
        btn_layout.addWidget(self.add_sender_btn)
        btn_layout.addWidget(self.edit_sender_btn)
        btn_layout.addWidget(self.remove_sender_btn)

        sender_layout.addLayout(btn_layout)
        sender_group.setLayout(sender_layout)
        main_layout.addWidget(sender_group)

        # Ignore Files group
        ignore_group = QtWidgets.QGroupBox("Ignore Files")
        ignore_layout = QtWidgets.QVBoxLayout()

        self.ignore_table = QtWidgets.QTableWidget()
        self.ignore_table.setColumnCount(1)
        self.ignore_table.setHorizontalHeaderLabels(["Filename or Pattern"])
        self.ignore_table.horizontalHeader().setStretchLastSection(True)
        self.ignore_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.ignore_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        ignore_layout.addWidget(self.ignore_table)

        self.populate_ignore_table()

        ignore_btn_layout = QtWidgets.QHBoxLayout()
        self.add_ignore_btn = QtWidgets.QPushButton("Add")
        self.edit_ignore_btn = QtWidgets.QPushButton("Edit")
        self.remove_ignore_btn = QtWidgets.QPushButton("Remove")
        self.add_ignore_btn.clicked.connect(self.add_ignore_item)
        self.edit_ignore_btn.clicked.connect(self.edit_ignore_item)
        self.remove_ignore_btn.clicked.connect(self.remove_ignore_item)
        ignore_btn_layout.addWidget(self.add_ignore_btn)
        ignore_btn_layout.addWidget(self.edit_ignore_btn)
        ignore_btn_layout.addWidget(self.remove_ignore_btn)

        ignore_layout.addLayout(ignore_btn_layout)
        ignore_group.setLayout(ignore_layout)
        main_layout.addWidget(ignore_group)

        # Process Invoices button
        self.process_invoices_btn = QtWidgets.QPushButton("Process Invoices")
        self.process_invoices_btn.clicked.connect(self.process_invoices_action)
        main_layout.addWidget(self.process_invoices_btn)

        self.status_label = QtWidgets.QLabel("Status: Idle")
        main_layout.addWidget(self.status_label)

        self.setLayout(main_layout)

    def load_json(self, path):
        if not os.path.exists(path):
            return None
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return None
    
    def save_json(self, path, data):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to save {path}:\n{e}")

    def populate_sender_table(self):
        self.sender_table.setRowCount(len(self.invoices_config_data))
        for row, item in enumerate(self.invoices_config_data):
            email = item.get("sender_email", "")
            folder_name = item.get("folder_name", "")
            file_name = item.get("file_name", "")
            month_offset = str(item.get("month_offset", 0))
            day_offset = str(item.get("day_offset", 0))

            self.sender_table.setItem(row, 0, QtWidgets.QTableWidgetItem(email))
            self.sender_table.setItem(row, 1, QtWidgets.QTableWidgetItem(folder_name))
            self.sender_table.setItem(row, 2, QtWidgets.QTableWidgetItem(file_name))
            self.sender_table.setItem(row, 3, QtWidgets.QTableWidgetItem(month_offset))
            self.sender_table.setItem(row, 4, QtWidgets.QTableWidgetItem(day_offset))

    def populate_ignore_table(self):
        self.ignore_table.setRowCount(len(self.ignore_files_data))
        for row, pattern in enumerate(self.ignore_files_data):
            self.ignore_table.setItem(row, 0, QtWidgets.QTableWidgetItem(pattern))

    def add_sender_config(self):
        dialog = SenderConfigDialog(self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            new_data = dialog.get_data()
            self.invoices_config_data.append(new_data)
            self.save_json(self.invoices_config_path, self.invoices_config_data)
            self.populate_sender_table()

    def edit_sender_config(self):
        row = self.sender_table.currentRow()
        if row < 0:
            return
        current_item = self.invoices_config_data[row]
        dialog = SenderConfigDialog(self, current_item)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            updated_data = dialog.get_data()
            self.invoices_config_data[row] = updated_data
            self.save_json(self.invoices_config_path, self.invoices_config_data)
            self.populate_sender_table()

    def remove_sender_config(self):
        row = self.sender_table.currentRow()
        if row < 0:
            return
        confirm = QtWidgets.QMessageBox.question(
            self, "Confirm Remove",
            "Remove this configuration?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if confirm == QtWidgets.QMessageBox.Yes:
            del self.invoices_config_data[row]
            self.save_json(self.invoices_config_path, self.invoices_config_data)
            self.populate_sender_table()

    def add_ignore_item(self):
        text, ok = QtWidgets.QInputDialog.getText(self, "Add Ignore Pattern", "Pattern:")
        if ok and text.strip():
            self.ignore_files_data.append(text.strip())
            self.save_json(self.invoices_ignore_path, self.ignore_files_data)
            self.populate_ignore_table()

    def edit_ignore_item(self):
        row = self.ignore_table.currentRow()
        if row < 0:
            return
        current_value = self.ignore_files_data[row]
        text, ok = QtWidgets.QInputDialog.getText(self, "Edit Ignore Pattern", "Pattern:", text=current_value)
        if ok and text.strip():
            self.ignore_files_data[row] = text.strip()
            self.save_json(self.invoices_ignore_path, self.ignore_files_data)
            self.populate_ignore_table()

    def remove_ignore_item(self):
        row = self.ignore_table.currentRow()
        if row < 0:
            return
        confirm = QtWidgets.QMessageBox.question(
            self, "Confirm Remove",
            "Remove this pattern?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if confirm == QtWidgets.QMessageBox.Yes:
            del self.ignore_files_data[row]
            self.save_json(self.invoices_ignore_path, self.ignore_files_data)
            self.populate_ignore_table()

    def process_invoices_action(self):
        """Invoke process_files(), capturing print output."""
        import io
        import sys

        from __main__ import process_files

        self.status_label.setText("Status: Processing...")
        QtWidgets.QApplication.processEvents()

        old_stdout = sys.stdout
        redirected_output = io.StringIO()
        sys.stdout = redirected_output

        try:
            process_files()
        except Exception as e:
            print(f"Error while processing files: {e}")

        sys.stdout = old_stdout
        output_text = redirected_output.getvalue()
        redirected_output.close()

        self.status_label.setText("Status: Idle")

        if output_text.strip():
            QtWidgets.QMessageBox.information(self, "Process Invoices", output_text)
        else:
            QtWidgets.QMessageBox.information(self, "Process Invoices", "Processing completed. No messages.")

################################################################################
#                           SENDER CONFIG DIALOG                               #
################################################################################

class SenderConfigDialog(QtWidgets.QDialog):
    """Dialog to add or edit sender config entries."""
    def __init__(self, parent=None, data=None):
        super(SenderConfigDialog, self).__init__(parent)
        self.setWindowTitle("Sender Configuration")
        self.setFixedSize(300, 250)

        self.data = data if data else {}

        layout = QtWidgets.QVBoxLayout()

        # Sender Email
        self.email_edit = QtWidgets.QLineEdit(self.data.get("sender_email", ""))
        layout.addWidget(QtWidgets.QLabel("Sender Email:"))
        layout.addWidget(self.email_edit)

        # Folder Name
        self.folder_edit = QtWidgets.QLineEdit(self.data.get("folder_name", ""))
        layout.addWidget(QtWidgets.QLabel("Folder Name:"))
        layout.addWidget(self.folder_edit)

        # File Name
        self.file_edit = QtWidgets.QLineEdit(self.data.get("file_name", ""))
        layout.addWidget(QtWidgets.QLabel("File Name:"))
        layout.addWidget(self.file_edit)

        # Month Offset
        self.month_offset_edit = QtWidgets.QSpinBox()
        self.month_offset_edit.setRange(-12, 12)
        self.month_offset_edit.setValue(int(self.data.get("month_offset", 0)))
        layout.addWidget(QtWidgets.QLabel("Month Offset:"))
        layout.addWidget(self.month_offset_edit)

        # Day Offset
        self.day_offset_edit = QtWidgets.QSpinBox()
        self.day_offset_edit.setRange(-31, 31)
        self.day_offset_edit.setValue(int(self.data.get("day_offset", 0)))
        layout.addWidget(QtWidgets.QLabel("Day Offset:"))
        layout.addWidget(self.day_offset_edit)

        # Dialog Buttons
        btn_layout = QtWidgets.QHBoxLayout()
        ok_btn = QtWidgets.QPushButton("OK")
        cancel_btn = QtWidgets.QPushButton("Cancel")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def get_data(self):
        return {
            "sender_email": self.email_edit.text().strip(),
            "folder_name": self.folder_edit.text().strip(),
            "file_name": self.file_edit.text().strip(),
            "month_offset": self.month_offset_edit.value(),
            "day_offset": self.day_offset_edit.value(),
        }

################################################################################
#                                    MAIN                                      #
################################################################################

def main():
    """Initialize the application and tray icon."""
    try:
        logging.basicConfig(level=logging.DEBUG)
        
        config = load_config(r'D:\Sync\Businesses\Worms Direct\Scripts\Downloading Attachments\config.ini')
        
        app = QtWidgets.QApplication(sys.argv)
        app.setQuitOnLastWindowClosed(False)
        
        # Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'icon.png')
        if not os.path.exists(icon_path):
            icon = QtGui.QIcon.fromTheme("mail-message-new")
            logging.warning("Custom icon not found; using default.")
        else:
            icon = QtGui.QIcon(icon_path)
            logging.info("Custom icon loaded.")
        
        tray = EmailAttachmentDownloader(icon, config)
        tray.show()
        
        sys.exit(app.exec_())
    except Exception as e:
        logging.exception("Unhandled exception in main().")
        print(f"An error occurred: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
