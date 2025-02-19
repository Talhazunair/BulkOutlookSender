import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QLineEdit, QTextEdit, QPushButton, 
                           QFileDialog, QProgressBar, QMessageBox, QTableWidget,
                           QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
from selenium.webdriver import ChromeOptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from undetected_chromedriver import Chrome
import pyperclip
import pyautogui
from selenium_stealth import stealth
import chromedriver_autoinstaller
import time
import random

class AutomationWorker(QThread):
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, excel_data, subject, body, attachment):
        super().__init__()
        print("Installing chromedriver...")
        chromedriver_autoinstaller.install()
        print("Chromedriver installed successfully.")
        self.excel_data = excel_data
        self.subject = subject
        self.body = body
        self.attachment = os.path.normpath(attachment)  # Normalize path
        self.driver = None

    def get_random_user_agent(self):
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0",
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        ]
        return random.choice(user_agents)

    def driver_setup(self):
        options = ChromeOptions()
        options.add_argument(f"user-agent={self.get_random_user_agent()}")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        self.driver = Chrome(options=options)
        stealth(
            self.driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
        )
        return self.driver

    def login(self, email, password):
        try:
            url = "https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=171&ct=1738948928&rver=7.5.2211.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26cobrandid%3dab0455a0-8d03-46b9-b18b-df2f57b9e44c"
            self.driver.get(url)
            
            email_input = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.NAME, "loginfmt")))
            email_input.send_keys(email)
            email_input.send_keys(Keys.RETURN)

            password_input = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.NAME, "passwd")))
            password_input.send_keys(password)
            password_input.send_keys(Keys.RETURN)
            
            time.sleep(5)

            no_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "declineButton")))
            no_button.click()

            time.sleep(5)
            return True
        except Exception as e:
            self.error.emit(f"Login failed for {email}: {str(e)}")
            return False

    def send_mail(self, to):
        try:
            new_email_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="114-group"]/div/div[1]/div/div/span/button[1]')))
            new_email_button.click()

            to_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@aria-label='To' and @role='textbox']")))
            to_field.click()
            pyperclip.copy(to)
            time.sleep(2)
            to_field.send_keys(Keys.CONTROL, 'v')

            subject_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Add a subject']")))
            subject_field.click()
            subject_field.send_keys(self.subject)

            body_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Message body, press Alt+F10 to exit']")))
            body_field.click()
            body_field.send_keys(self.body)

            insert_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="5"]')))
            insert_button.click()
            time.sleep(3)
            attach_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Attach file']")))
            attach_button.click()
            time.sleep(1)
            browse_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Browse this computer']")))
            browse_button.click()
            time.sleep(3)
            
            # Handle file path correctly
            normalized_path = os.path.normpath(self.attachment)
            pyperclip.copy(normalized_path)
            time.sleep(1)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(13)
            
            send_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Send']")))
            send_button.click()

            return True
        except Exception as e:
            self.error.emit(f"Error sending email: {str(e)}")
            return False

    def run(self):
        try:
            total_accounts = len(self.excel_data)
            for index, row in self.excel_data.iterrows():
                progress = int((index / total_accounts) * 100)
                self.progress_updated.emit(progress, f"Processing account: {row['email']}")
                
                self.driver_setup()
                if self.login(row['email'], row['password']):
                    if self.send_mail(row['to']):
                        self.progress_updated.emit(progress, f"Email sent successfully from {row['email']}")
                    
                if self.driver:
                    self.driver.quit()
                    self.driver = None
                
                delay = random.uniform(5, 10)
                time.sleep(delay)
                
            self.progress_updated.emit(100, "All accounts processed")
            self.finished.emit()
            
        except Exception as e:
            self.error.emit(f"Error in automation: {str(e)}")
        finally:
            if self.driver:
                self.driver.quit()

class OutlookAutomationGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_data = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Outlook Email Automation')
        self.setGeometry(100, 100, 800, 600)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Excel file selection
        excel_layout = QHBoxLayout()
        self.excel_path = QLineEdit()
        excel_button = QPushButton('Select Excel File')
        excel_button.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(QLabel('Excel File:'))
        excel_layout.addWidget(self.excel_path)
        excel_layout.addWidget(excel_button)
        layout.addLayout(excel_layout)

        # Data preview table
        self.table = QTableWidget()
        self.table.setVisible(False)
        layout.addWidget(self.table)

        # Email details
        self.subject = QLineEdit()
        self.body = QTextEdit()
        self.attachment_path = QLineEdit()
        attachment_button = QPushButton('Select Attachment')
        attachment_button.clicked.connect(self.select_attachment)

        form_layout = QVBoxLayout()
        form_layout.addWidget(QLabel('Subject:'))
        form_layout.addWidget(self.subject)
        form_layout.addWidget(QLabel('Body:'))
        form_layout.addWidget(self.body)
        
        attachment_layout = QHBoxLayout()
        attachment_layout.addWidget(QLabel('Attachment:'))
        attachment_layout.addWidget(self.attachment_path)
        attachment_layout.addWidget(attachment_button)
        form_layout.addLayout(attachment_layout)
        
        layout.addLayout(form_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_label = QLabel('Ready')
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.progress_label)

        # Start button
        self.start_button = QPushButton('Start Automation')
        self.start_button.clicked.connect(self.start_automation)
        self.start_button.setEnabled(False)
        layout.addWidget(self.start_button)

    def select_excel_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Excel File', '', 'Excel Files (*.xlsx *.xls)')
        if file_name:
            normalized_path = os.path.normpath(file_name)
            self.excel_path.setText(normalized_path)
            self.load_excel_data(normalized_path)

    def load_excel_data(self, file_name):
        try:
            self.excel_data = pd.read_excel(file_name)
            required_columns = ['email', 'password', 'to']
            missing_columns = [col for col in required_columns if col not in self.excel_data.columns]
            
            if missing_columns:
                QMessageBox.critical(self, 'Error', f"Missing required columns: {', '.join(missing_columns)}")
                return

            # Update table
            self.table.setVisible(True)
            self.table.setRowCount(len(self.excel_data))
            self.table.setColumnCount(len(self.excel_data.columns))
            self.table.setHorizontalHeaderLabels(self.excel_data.columns)

            for i in range(len(self.excel_data)):
                for j in range(len(self.excel_data.columns)):
                    item = QTableWidgetItem(str(self.excel_data.iloc[i, j]))
                    self.table.setItem(i, j, item)

            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            self.start_button.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, 'Error', f"Error reading Excel file: {str(e)}")

    def select_attachment(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Attachment', '', 'All Files (*.*)')
        if file_name:
            normalized_path = os.path.normpath(file_name)
            self.attachment_path.setText(normalized_path)

    def start_automation(self):
        if not all([self.excel_data is not None, 
                   self.subject.text(), 
                   self.body.toPlainText(), 
                   self.attachment_path.text()]):
            QMessageBox.warning(self, 'Warning', 'Please fill in all fields')
            return

        self.start_button.setEnabled(False)
        self.worker = AutomationWorker(
            self.excel_data,
            self.subject.text(),
            self.body.toPlainText(),
            self.attachment_path.text()
        )
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.finished.connect(self.automation_finished)
        self.worker.error.connect(self.show_error)
        self.worker.start()

    def update_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)

    def automation_finished(self):
        self.start_button.setEnabled(True)
        QMessageBox.information(self, 'Complete', 'Automation completed successfully!')

    def show_error(self, message):
        QMessageBox.critical(self, 'Error', message)

def main():
    app = QApplication(sys.argv)
    ex = OutlookAutomationGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()