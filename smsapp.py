import sys
import pandas as pd
import smtplib
import json
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout, QWidget, QTextEdit, QCheckBox, QFileDialog, QProgressBar, QFormLayout, QListWidget, QListWidgetItem, QInputDialog, QMessageBox, QRadioButton, QButtonGroup, QComboBox, QTabWidget, QScrollArea
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Hardcoded license key
LICENSE_KEY = 'oladdev'

# Helper Function to Check Spammy Words
def check_spam(text):
    spam_words = ["free", "win", "cash", "prize", "winner", "guaranteed"]
    return any(word in text.lower() for word in spam_words)

# Function to Check License
def check_license():
    license_input, ok = QInputDialog.getText(None, 'License Check', 'Enter license key:', QLineEdit.Password)
    if not ok or license_input != LICENSE_KEY:
        QMessageBox.critical(None, 'License Error', 'Invalid license key. Exiting application.')
        sys.exit(1)

# Function to format and append SMS gateway
def append_gateway(number, country, gateway):
    # Remove country code prefix based on the selected country
    if country == 'United States' and number.startswith('+1'):
        number = number[2:]  # Remove +1
    elif country == 'United Kingdom' and number.startswith('+44'):
        number = number[3:]  # Remove +44
    elif country == 'Canada' and number.startswith('+1'):
        number = number[2:]  # Remove +1
    elif country == 'Australia' and number.startswith('+61'):
        number = number[3:]  # Remove +61
    
    # Remove any spaces, hyphens, or plus signs
    number = number.replace('+', '').replace('-', '').replace(' ', '').strip()

    # Append the correct SMS gateway
    if country == 'United States':
        if gateway == 'AT&T':
            return f"{number}@txt.att.net"
        elif gateway == 'T-Mobile':
            return f"{number}@tmomail.net"
        elif gateway == 'Verizon':
            return f"{number}@vtext.com"
        elif gateway == 'Sprint':
            return f"{number}@messaging.sprintpcs.com"
    elif country == 'United Kingdom':
        if gateway == 'Vodafone UK':
            return f"{number}@vodafone.net"
        elif gateway == 'O2':
            return f"{number}@o2.co.uk"
    elif country == 'Canada':
        if gateway == 'Rogers':
            return f"{number}@pcs.rogers.com"
        elif gateway == 'Bell':
            return f"{number}@txt.bell.ca"
        elif gateway == 'Telus':
            return f"{number}@msg.telus.com"
        elif gateway == 'Fido':
            return f"{number}@fido.ca"
    elif country == 'Australia':
        if gateway == 'Telstra':
            return f"{number}@sms.telstra.com"
        elif gateway == 'Optus':
            return f"{number}@optusmobile.com.au"
        elif gateway == 'Vodafone AU':
            return f"{number}@vfa.com.au"
    return None

# Worker Thread for Sending SMS/Email
class MessageSenderThread(QThread):
    progress = pyqtSignal(int)
    status_update = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    log_update = pyqtSignal(dict)

    def __init__(self, smtp_details, message_text, leads, rotate_count, subject, speed_value, speed_unit, attachments, parent=None):
        super().__init__(parent)
        self.smtp_details = smtp_details
        self.message_text = message_text
        self.leads = leads
        self.rotate_count = rotate_count
        self.subject = subject
        self.speed_value = speed_value
        self.speed_unit = speed_unit
        self.attachments = attachments
        self.paused = False

    def run(self):
        success_count = 0
        failed_count = 0
        total_leads = len(self.leads)
        index = 0

        # Calculate sending limits based on speed
        if self.speed_unit == 'hour':
            send_limit = self.speed_value
            sleep_time = 3600  # 1 hour
        else:
            send_limit = self.speed_value
            sleep_time = 60  # 1 minute

        while index < total_leads:
            if self.isInterruptionRequested():
                break

            if self.paused:
                self.msleep(1000)  # Sleep for a while if paused
                continue

            lead = self.leads[index]
            smtp_index = (index // self.rotate_count) % len(self.smtp_details)
            smtp = self.smtp_details[smtp_index]

            try:
                with smtplib.SMTP(smtp['host'], smtp['port']) as server:
                    server.starttls()
                    server.login(smtp['username'], smtp['password'])

                    msg = MIMEMultipart()
                    msg['From'] = smtp['from_email']
                    msg['To'] = lead
                    msg['Subject'] = self.subject

                    msg.attach(MIMEText(self.message_text, 'html'))

                    # Attach files
                    for attachment in self.attachments:
                        with open(attachment, "rb") as file:
                            part = MIMEApplication(file.read(), Name=os.path.basename(attachment))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
                        msg.attach(part)

                    server.sendmail(smtp['from_email'], lead, msg.as_string())

                success_count += 1
                status = f"Sent to {lead}: Success with {smtp['username']} ({smtp['sender_name']})"
                self.status_update.emit(status)
                self.log_update.emit({
                    'recipient': lead,
                    'status': 'Success',
                    'smtp': f"{smtp['username']} ({smtp['sender_name']})",
                    'message': status
                })
            except Exception as e:
                status = f"Failed to send to {lead}: {e} with {smtp['username']} ({smtp['sender_name']})"
                self.status_update.emit(status)
                self.log_update.emit({
                    'recipient': lead,
                    'status': 'Failed',
                    'smtp': f"{smtp['username']} ({smtp['sender_name']})",
                    'message': status
                })
                failed_count += 1

            index += 1
            self.progress.emit(int((index) / total_leads * 100))

            # Check if we have reached the send limit for this time unit
            if index % send_limit == 0:
                self.status_update.emit(f"Reached send limit of {send_limit}. Waiting for the next time slot...")
                self.msleep(sleep_time * 1000)  # Sleep for 1 hour or 1 minute, based on the selected speed

        self.finished.emit(True, f"Completed: {success_count} sent, {failed_count} failed")

    def pause(self):
        self.paused = True

    def resume(self):
        self.paused = False

class SMSApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SMTP to SMS/Email Sender")
        self.setGeometry(100, 100, 1200, 600)
        self.smtp_details = self.load_smtp_details()  # Load saved SMTP details
        self.initUI()

    def initUI(self):
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        # Create tabs
        self.smtp_tab = QWidget()
        self.message_tab = QWidget()
        self.logs_tab = QWidget()

        self.tab_widget.addTab(self.smtp_tab, "SMTP Settings")
        self.tab_widget.addTab(self.message_tab, "Message Sender")
        self.tab_widget.addTab(self.logs_tab, "Email Logs")

        self.setup_smtp_tab()
        self.setup_message_tab()
        self.setup_logs_tab()

        # Footer
        footer = QLabel('<a href="https://bit.ly/hiolad">Developed by Olad Synergy Solutions</a>')
        footer.setOpenExternalLinks(True)
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("font-size: 12px; color: gray;")
        self.statusBar().addPermanentWidget(footer)

    def setup_smtp_tab(self):
        layout = QVBoxLayout()

        # Form for SMTP Details
        smtp_form_layout = QFormLayout()

        self.smtp_host = QLineEdit()
        self.smtp_port = QLineEdit()
        self.smtp_username = QLineEdit()
        self.smtp_password = QLineEdit()
        self.from_email = QLineEdit()
        self.sender_name = QLineEdit()

        smtp_form_layout.addRow(QLabel('SMTP Host:'), self.smtp_host)
        smtp_form_layout.addRow(QLabel('SMTP Port:'), self.smtp_port)
        smtp_form_layout.addRow(QLabel('SMTP Username:'), self.smtp_username)
        smtp_form_layout.addRow(QLabel('SMTP Password:'), self.smtp_password)
        smtp_form_layout.addRow(QLabel('From Email:'), self.from_email)
        smtp_form_layout.addRow(QLabel('Sender Name:'), self.sender_name)

        self.test_add_button = QPushButton('Test AND Add SMTP')
        self.remove_smtp_button = QPushButton('Remove Selected SMTP')
        self.remove_smtp_button.setEnabled(False)

        self.test_add_button.setStyleSheet("background-color: blue; color: white;")
        self.remove_smtp_button.setStyleSheet("background-color: red; color: white;")

        smtp_form_layout.addWidget(self.test_add_button)
        smtp_form_layout.addWidget(self.remove_smtp_button)

        layout.addLayout(smtp_form_layout)

        # ListWidget for SMTP Details
        self.smtp_list_widget = QListWidget()
        self.smtp_list_widget.itemSelectionChanged.connect(self.on_smtp_selection_changed)
        for smtp in self.smtp_details:
            self.smtp_list_widget.addItem(QListWidgetItem(f"{smtp['host']}:{smtp['port']} - {smtp['username']}"))
        layout.addWidget(self.smtp_list_widget)

        # Read-only box for SMTP Details
        self.smtp_display = QTextEdit()
        self.smtp_display.setPlaceholderText('Added SMTP details will appear here...')
        self.smtp_display.setReadOnly(True)
        layout.addWidget(self.smtp_display)

        self.smtp_tab.setLayout(layout)

        # Connections
        self.test_add_button.clicked.connect(self.test_and_add_smtp)
        self.remove_smtp_button.clicked.connect(self.remove_smtp)

        # Update the SMTP display
        self.update_smtp_display()

    def setup_message_tab(self):
        layout = QVBoxLayout()

        # Message content
        self.message_text_edit = QTextEdit()
        self.message_text_edit.setPlaceholderText('Enter SMS/EMAIL content here...')
        layout.addWidget(QLabel('Message Content:'))
        layout.addWidget(self.message_text_edit)

        # Subject
        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText('Enter Subject Here')
        layout.addWidget(QLabel('Subject:'))
        layout.addWidget(self.subject_edit)

        # Leads
        self.leads_text_edit = QTextEdit()
        self.leads_text_edit.setPlaceholderText('Enter leads here, one per line...')
        layout.addWidget(QLabel('Leads:'))
        layout.addWidget(self.leads_text_edit)

        # Upload leads button
        self.upload_button = QPushButton('Upload Leads')
        self.upload_button.clicked.connect(self.upload_leads)
        self.upload_button.setStyleSheet("background-color: blue; color: white;")
        layout.addWidget(self.upload_button)

        # Country and Gateway Selection
        country_gateway_layout = QHBoxLayout()
        self.country_combo = QComboBox()
        self.country_combo.addItems(['United States', 'United Kingdom', 'Canada', 'Australia'])
        self.gateway_combo = QComboBox()
        country_gateway_layout.addWidget(QLabel('Country:'))
        country_gateway_layout.addWidget(self.country_combo)
        country_gateway_layout.addWidget(QLabel('Gateway:'))
        country_gateway_layout.addWidget(self.gateway_combo)
        layout.addLayout(country_gateway_layout)

        # Button to append gateway to leads
        self.append_gateway_button = QPushButton('Append Gateway to Uploaded Leads')
        self.append_gateway_button.setStyleSheet("background-color: green; color: white;")
        layout.addWidget(self.append_gateway_button)

        # SMTP rotation
        rotation_layout = QHBoxLayout()
        self.rotate_count_checkbox = QCheckBox('Rotate SMTP after sending')
        self.rotate_count_checkbox.setChecked(True)
        self.rotate_count_spinbox = QLineEdit()
        self.rotate_count_spinbox.setText('1')
        rotation_layout.addWidget(self.rotate_count_checkbox)
        rotation_layout.addWidget(QLabel('SMTP Rotate Count:'))
        rotation_layout.addWidget(self.rotate_count_spinbox)
        layout.addLayout(rotation_layout)

        # Speed Selection
        speed_layout = QHBoxLayout()
        self.per_hour_radio = QRadioButton('Per Hour')
        self.per_minute_radio = QRadioButton('Per Minute')
        self.per_hour_radio.setChecked(True)
        self.speed_button_group = QButtonGroup()
        self.speed_button_group.addButton(self.per_hour_radio)
        self.speed_button_group.addButton(self.per_minute_radio)
        self.speed_value_input = QLineEdit()
        self.speed_value_input.setPlaceholderText('Enter speed value (e.g., 2000)')
        speed_layout.addWidget(QLabel('Select Speed:'))
        speed_layout.addWidget(self.per_hour_radio)
        speed_layout.addWidget(self.per_minute_radio)
        speed_layout.addWidget(self.speed_value_input)
        layout.addLayout(speed_layout)

        # File attachments
        self.attachment_list = QListWidget()
        layout.addWidget(QLabel('Attachments:'))
        layout.addWidget(self.attachment_list)
        attachment_buttons_layout = QHBoxLayout()
        self.add_attachment_button = QPushButton('Add Attachment')
        self.remove_attachment_button = QPushButton('Remove Attachment')
        self.add_attachment_button.clicked.connect(self.add_attachment)
        self.remove_attachment_button.clicked.connect(self.remove_attachment)
        attachment_buttons_layout.addWidget(self.add_attachment_button)
        attachment_buttons_layout.addWidget(self.remove_attachment_button)
        layout.addLayout(attachment_buttons_layout)

        # Progress bar and buttons
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        button_layout = QHBoxLayout()
        self.start_button = QPushButton('Start Sending')
        self.pause_button = QPushButton('Pause')
        self.stop_button = QPushButton('Stop')
        self.clear_button = QPushButton('Clear Progress')
        self.start_button.setStyleSheet("background-color: red; color: white;")
        self.pause_button.setStyleSheet("background-color: gold; color: black;")
        self.stop_button.setStyleSheet("background-color: red; color: white;")
        self.clear_button.setStyleSheet("background-color: gray; color: white;")
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.pause_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(self.clear_button)
        layout.addLayout(button_layout)

        self.progress_box = QTextEdit()
        self.progress_box.setPlaceholderText('Progress messages will appear here...')
        self.progress_box.setReadOnly(True)
        layout.addWidget(self.progress_box)

        self.message_tab.setLayout(layout)

        # Connections
        self.country_combo.currentIndexChanged.connect(self.update_gateway)
        self.append_gateway_button.clicked.connect(self.append_gateway_to_leads)
        self.start_button.clicked.connect(self.start_sending)
        self.pause_button.clicked.connect(self.toggle_pause_resume)
        self.stop_button.clicked.connect(self.stop_sending)
        self.clear_button.clicked.connect(self.clear_progress)

        self.update_gateway()  # Initialize the gateway combo box

    def setup_logs_tab(self):
        layout = QVBoxLayout()

        self.logs_text_edit = QTextEdit()
        self.logs_text_edit.setReadOnly(True)
        self.logs_text_edit.setPlaceholderText('Email logs will appear here...')

        layout.addWidget(self.logs_text_edit)

        self.logs_tab.setLayout(layout)

    def load_smtp_details(self):
        try:
            with open('smtp_details.json', 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            return []

    def save_smtp_details(self):
        with open('smtp_details.json', 'w') as f:
            json.dump(self.smtp_details, f)

    def test_and_add_smtp(self):
        host = self.smtp_host.text()
        port = int(self.smtp_port.text())
        username = self.smtp_username.text()
        password = self.smtp_password.text()
        from_email = self.from_email.text()
        sender_name = self.sender_name.text()

        if not (host and port and username and password and from_email and sender_name):
            QMessageBox.warning(self, 'Input Error', 'All SMTP fields must be filled.')
            return

        try:
            with smtplib.SMTP(host, port) as server:
                server.starttls()
                server.login(username, password)
                test_msg = MIMEText('This is a test email.', 'html')
                test_msg['From'] = from_email
                test_msg['To'] = from_email
                test_msg['Subject'] = 'Test Email'
                server.sendmail(from_email, from_email, test_msg.as_string())
            
            new_smtp = {
                'host': host,
                'port': port,
                'username': username,
                'password': password,
                'from_email': from_email,
                'sender_name': sender_name
            }
            self.smtp_details.append(new_smtp)

            self.smtp_list_widget.addItem(QListWidgetItem(f"{host}:{port} - {username}"))
            self.save_smtp_details()
            self.update_smtp_display()

            QMessageBox.information(self, 'Success', 'SMTP configuration is correct and added.')
        except Exception as e:
            QMessageBox.critical(self, 'SMTP Error', f'Error testing SMTP configuration: {e}')

    def remove_smtp(self):
        current_item = self.smtp_list_widget.currentItem()
        if current_item:
            row = self.smtp_list_widget.row(current_item)
            del self.smtp_details[row]
            self.smtp_list_widget.takeItem(row)
            self.save_smtp_details()
            self.update_smtp_display()
            self.remove_smtp_button.setEnabled(False)

    def on_smtp_selection_changed(self):
        self.remove_smtp_button.setEnabled(bool(self.smtp_list_widget.selectedItems()))

    def update_smtp_display(self):
        self.smtp_display.clear()
        for smtp in self.smtp_details:
            self.smtp_display.append(f"Host: {smtp['host']}")
            self.smtp_display.append(f"Port: {smtp['port']}")
            self.smtp_display.append(f"Username: {smtp['username']}")
            self.smtp_display.append(f"From Email: {smtp['from_email']}")
            self.smtp_display.append(f"Sender Name: {smtp['sender_name']}")
            self.smtp_display.append("-" * 30)

    def upload_leads(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open Leads File', '', 'Text Files (*.txt);;All Files (*)')
        if file_path:
            with open(file_path, 'r') as file:
                leads = file.readlines()
            self.leads_text_edit.setText(''.join(leads))

    def update_gateway(self):
        country = self.country_combo.currentText()
        self.gateway_combo.clear()
        if country == 'United States':
            self.gateway_combo.addItems(['AT&T', 'T-Mobile', 'Verizon', 'Sprint'])
        elif country == 'United Kingdom':
            self.gateway_combo.addItems(['Vodafone UK', 'O2'])
        elif country == 'Canada':
            self.gateway_combo.addItems(['Rogers', 'Bell', 'Telus', 'Fido'])
        elif country == 'Australia':
            self.gateway_combo.addItems(['Telstra', 'Optus', 'Vodafone AU'])

    def append_gateway_to_leads(self):
        leads = self.leads_text_edit.toPlainText().splitlines()
        country = self.country_combo.currentText()
        gateway = self.gateway_combo.currentText()

        appended_leads = []
        for lead in leads:
            appended_lead = append_gateway(lead, country, gateway)
            if appended_lead:
                appended_leads.append(appended_lead)

        self.leads_text_edit.setText('\n'.join(appended_leads))

    def add_attachment(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Attachment', '', 'All Files (*)')
        if file_path:
            self.attachment_list.addItem(file_path)

    def remove_attachment(self):
        current_item = self.attachment_list.currentItem()
        if current_item:
            self.attachment_list.takeItem(self.attachment_list.row(current_item))

    def start_sending(self):
        if not self.smtp_details:
            QMessageBox.warning(self, 'No SMTP Details', 'Please add SMTP details before starting.')
            return
        
        if check_spam(self.message_text_edit.toPlainText()):
            QMessageBox.warning(self, 'Spam Alert', 'The content contains spammy words.')
            return

        # Get speed settings
        if self.per_hour_radio.isChecked():
            speed_unit = 'hour'
        else:
            speed_unit = 'minute'

        try:
            speed_value = int(self.speed_value_input.text())
        except ValueError:
            QMessageBox.warning(self, 'Input Error', 'Speed value must be a valid integer.')
            return

        attachments = [self.attachment_list.item(i).text() for i in range(self.attachment_list.count())]

        self.sender_thread = MessageSenderThread(
            smtp_details=self.smtp_details,
            message_text=self.message_text_edit.toPlainText(),
            leads=self.leads_text_edit.toPlainText().splitlines(),
            rotate_count=int(self.rotate_count_spinbox.text()),
            subject=self.subject_edit.text(),
            speed_value=speed_value,
            speed_unit=speed_unit,
            attachments=attachments
        )

        self.sender_thread.progress.connect(self.progress_bar.setValue)
        self.sender_thread.status_update.connect(self.progress_box.append)
        self.sender_thread.finished.connect(self.on_sending_finished)
        self.sender_thread.log_update.connect(self.update_logs)
        self.sender_thread.start()

        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.stop_button.setEnabled(True)

    def toggle_pause_resume(self):
        if self.sender_thread.paused:
            self.sender_thread.resume()
            self.pause_button.setText('Pause')
        else:
            self.sender_thread.pause()
            self.pause_button.setText('Resume')

    def stop_sending(self):
        if hasattr(self, 'sender_thread'):
            self.sender_thread.requestInterruption()
            self.sender_thread.wait()

        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)

    def clear_progress(self):
        self.progress_box.clear()
        self.progress_bar.setValue(0)

    def on_sending_finished(self, success, message):
        QMessageBox.information(self, 'Sending Completed', message)
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)

    def update_logs(self, log_entry):
        log_text = f"Recipient: {log_entry['recipient']}\n"
        log_text += f"Status: {log_entry['status']}\n"
        log_text += f"SMTP: {log_entry['smtp']}\n"
        log_text += f"Message: {log_entry['message']}\n"
        log_text += "-" * 50 + "\n"
        self.logs_text_edit.append(log_text)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    check_license()
    window = SMSApp()
    window.show()
    sys.exit(app.exec_())