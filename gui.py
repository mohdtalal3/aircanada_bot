import sys
import os
import csv
import time
import re
import random
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QPushButton, QLabel, QTextEdit, 
                              QFileDialog, QProgressBar, QGroupBox, QLineEdit)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont, QTextCursor
from seleniumbase import SB


class WorkerThread(QThread):
    """Worker thread to process CSV entries without freezing the GUI"""
    log_signal = pyqtSignal(str, str)  # message, level (info/success/error/warning)
    progress_signal = pyqtSignal(int, int)  # current, total
    finished_signal = pyqtSignal()
    
    def __init__(self, csv_file, processed_file):
        super().__init__()
        self.csv_file = csv_file
        self.processed_file = processed_file
        self.is_running = True
        self.extension_dir = os.path.join(os.getcwd(), 'extension')
        
    def stop(self):
        """Stop the worker thread"""
        self.is_running = False
        
    def log(self, message, level="info"):
        """Emit log message"""
        self.log_signal.emit(message, level)
        
    def read_csv(self, file_path):
        """Read CSV file and return list of dictionaries"""
        data_list = []
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                if row.get('Email') and row.get('Email').strip():
                    data_list.append(row)
        return data_list
    
    def get_processed_emails(self):
        """Get list of already processed emails"""
        if not os.path.exists(self.processed_file):
            return set()
        
        processed_emails = set()
        with open(self.processed_file, 'r', encoding='utf-8-sig') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                if row.get('Email'):
                    processed_emails.add(row.get('Email').strip())
        return processed_emails
    
    def save_processed_entry(self, row_data, aeroplan_number=None):
        """Save processed entry to CSV file"""
        file_exists = os.path.exists(self.processed_file)
        
        with open(self.processed_file, 'a', encoding='utf-8', newline='') as file:
            fieldnames = list(row_data.keys())
            if 'Processed_Date' not in fieldnames:
                fieldnames.append('Processed_Date')
            if 'Aeroplan number' not in fieldnames:
                fieldnames.append('Aeroplan number')
                
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            
            if not file_exists:
                writer.writeheader()
            
            row_data['Processed_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            if aeroplan_number:
                row_data['Aeroplan number'] = aeroplan_number
            else:
                row_data['Aeroplan number'] = 'Failed to extract'
                
            writer.writerow(row_data)
    
    def remove_processed_row_from_csv(self, email):
        """Remove processed row from original CSV"""
        try:
            # Read all rows
            rows = []
            with open(self.csv_file, 'r', encoding='utf-8-sig') as file:
                csv_reader = csv.DictReader(file)
                fieldnames = csv_reader.fieldnames
                for row in csv_reader:
                    if row.get('Email', '').strip() != email:
                        rows.append(row)
            
            # Write back without the processed row
            with open(self.csv_file, 'w', encoding='utf-8-sig', newline='') as file:
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
                
            self.log(f"✓ Removed {email} from original CSV", "success")
        except Exception as e:
            self.log(f"✗ Failed to remove row from CSV: {str(e)}", "error")
    
    def format_proxy(self, proxy_string):
        """Convert proxy format from IP:PORT:USERNAME:PASSWORD to USERNAME:PASSWORD@IP:PORT"""
        if not proxy_string or not proxy_string.strip():
            return None
        
        parts = proxy_string.strip().split(':')
        if len(parts) == 4:
            ip, port, username, password = parts
            return f"{username}:{password}@{ip}:{port}"
        return proxy_string  # Return as-is if already in correct format
    
    def fill_form(self, sb, data):
        """Fill the form with data from CSV row"""
        try:
            sb.maximize_window()
            self.log("Opening registration page...", "info")
            sb.open("https://www.aircanada.com/aeroplan/member/enrolment")
            time.sleep(random.uniform(1, 3))
            
            # Fill email
            self.log("Filling email...", "info")
            sb.js_click("#emailFocus", scroll=True, timeout=60)
            sb.type("#emailFocus", data.get("Email"))
            time.sleep(random.uniform(1, 3))

            # Fill password
            sb.js_click("#pwd", scroll=True, timeout=15)
            sb.type("#pwd", data.get("Password"))
            time.sleep(random.uniform(1, 3))

            # Click checkbox
            sb.js_click('input[id="checkBox-input"]', timeout=15)
            time.sleep(random.uniform(1, 3))

            # Click continue button
            self.log("Submitting credentials...", "info")
            sb.js_click("button[data-analytics-val*='continue']", timeout=15)
            time.sleep(random.uniform(2, 4))

            # Fill personal details
            self.log("Filling personal details...", "info")
            sb.js_click("input[name='firstName']", scroll=True, timeout=15)
            sb.type("input[name='firstName']", data.get("First name"))
            time.sleep(random.uniform(1, 3))

            sb.js_click("input[name='lastName']", scroll=True, timeout=15)
            sb.type("input[name='lastName']", data.get("Last name"))
            time.sleep(random.uniform(1, 3))

            # Select gender
            sb.js_click('mat-select[formcontrolname="gender"]', scroll=True, timeout=15)
            time.sleep(random.uniform(1, 2))
            if data.get("Gender") == "Male":
                sb.js_click("//mat-option//span[text()=' Male ']", by="xpath", timeout=15)
            else:
                sb.js_click("//mat-option//span[text()=' Female ']", by="xpath", timeout=15)
            time.sleep(random.uniform(1, 3))

            # Select birthday
            self.log("Selecting birth date...", "info")
            sb.js_click("mat-select[formcontrolname='d']", scroll=True, timeout=15)
            time.sleep(random.uniform(1, 2))
            sb.js_click(f"//mat-option//span[text()=' {data.get("Day")} ']", by="xpath", timeout=15)
            time.sleep(random.uniform(1, 3))

            sb.js_click("mat-select[formcontrolname='m']", scroll=True, timeout=15)
            time.sleep(random.uniform(1, 2))
            sb.js_click(f"//mat-option//span[text()=' {data.get("Month")} ']", by="xpath", timeout=15)
            time.sleep(random.uniform(1, 3))

            sb.js_click("mat-select[formcontrolname='y']", scroll=True, timeout=15)
            time.sleep(random.uniform(1, 2))
            sb.js_click(f"//mat-option//span[text()=' {data.get("Year")} ']", by="xpath", timeout=15)
            time.sleep(random.uniform(1, 3))

            # Click next button
            sb.js_click("button[data-analytics-val*='continue']", timeout=5)
            time.sleep(random.uniform(2, 4))
            
            # Wait for CAPTCHA
            self.log("⚠️ Waiting for CAPTCHA to be solved...", "warning")
            sb.wait_for_element_visible('iframe[title="reCAPTCHA"]', timeout=30)
            sb.switch_to_frame('iframe[title="reCAPTCHA"]')
            sb.wait_for_element_present('span.recaptcha-checkbox[aria-checked="true"]', timeout=300)
            sb.switch_to_default_content()
            self.log("✓ CAPTCHA solved!", "success")
            time.sleep(random.uniform(1, 3))
            
            # Fill address
            self.log("Filling address details...", "info")
            sb.js_click('input[formcontrolname="addressLine1"]', scroll=True, timeout=15)
            sb.type('input[formcontrolname="addressLine1"]', data.get("Address"), timeout=15)
            time.sleep(random.uniform(1, 3))

            sb.js_click('input[formcontrolname="city"]', scroll=True, timeout=15)
            sb.type('input[formcontrolname="city"]', data.get("City"), timeout=15)
            time.sleep(random.uniform(1, 3))

            # Select country
            sb.js_click('mat-select[formcontrolname="country"]', scroll=True, timeout=15)
            time.sleep(random.uniform(1, 2))
            sb.js_click(f"//mat-option//span[text()=' {data.get("Country")} ']", by="xpath", timeout=15)
            time.sleep(random.uniform(1, 3))

            # Select province/state
            sb.js_click('mat-select[formcontrolname="state"]', scroll=True, timeout=15)
            time.sleep(random.uniform(1, 2))
            sb.js_click(f"//mat-option//span[text()=' {data.get("Province")} ']", by="xpath", timeout=15)
            time.sleep(random.uniform(1, 3))

            # Fill postal code
            sb.js_click('input[formcontrolname="zip"]', scroll=True, timeout=15)
            sb.type('input[formcontrolname="zip"]', data.get("Postal Code"), timeout=15)
            time.sleep(random.uniform(1, 3))
            
            # Fill phone number
            sb.js_click('input[formcontrolname="phoneNumber"]', scroll=True, timeout=15)
            sb.type('input[formcontrolname="phoneNumber"]', data.get("Phone number"), timeout=15)
            time.sleep(random.uniform(1, 3))

            # Click privacy policy checkbox
            sb.js_click('input[id="privacyPolicycheckBox"]', scroll=True, timeout=15)
            time.sleep(random.uniform(1, 3))

            # Submit form
            self.log("Submitting registration form...", "info")
            sb.js_click('button[data-analytics-val*="create my account"]', scroll=True, timeout=15)
            
            time.sleep(random.uniform(3, 5))
            
            # Extract Aeroplan number
            try:
                self.log("Extracting Aeroplan number...", "info")
                sb.wait_for_element_present('span.aeroplan-number', timeout=300)
                number_text = sb.get_text("span.aeroplan-number")
                aeroplan_number = re.sub(r"\D", "", number_text)
                self.log(f"🎉 Aeroplan Number: {aeroplan_number}", "success")
                time.sleep(random.uniform(1, 3))
                
                if aeroplan_number:
                    self.log("🎭 Starting Disney promotion registration...", "info")
                    sb.open("https://www.aircanada.com/ca/en/aco/home/book/special-offers/disney-promotion.html")
                    self.log("✓ Opened Disney promotion page", "success")
                    time.sleep(random.uniform(2, 4))

                    self.log("Waiting for Disney form to load...", "info")
                    sb.wait_for_element_present('input[name="ae-number"]', timeout=300)
                    time.sleep(random.uniform(1, 3))
                    
                    self.log("Entering Aeroplan number in Disney form...", "info")
                    sb.js_click('input[name="ae-number"]', scroll=True, timeout=15)
                    sb.type('input[name="ae-number"]', aeroplan_number, timeout=15)
                    self.log(f"✓ Entered Aeroplan number: {aeroplan_number}", "success")
                    time.sleep(random.uniform(1, 3))

                    self.log("Accepting agreement terms...", "info")
                    sb.js_click('input[name="agreement"]', scroll=True, timeout=15)
                    self.log("✓ Agreement checkbox checked", "success")
                    time.sleep(random.uniform(1, 3))

                    self.log("Checking profile confirmation...", "info")
                    sb.js_click('input[name="profilecheck"]', scroll=True, timeout=15)
                    self.log("✓ Profile check completed", "success")
                    time.sleep(random.uniform(1, 3))

                    self.log("Submitting Disney promotion form...", "info")
                    sb.js_click('input[type="submit"]', scroll=True, timeout=15)
                    time.sleep(random.uniform(3, 5))
                    self.log("🎉 Disney promotion registration completed!", "success")
                return aeroplan_number
            except Exception as e:
                self.log(f"✗ Could not extract Aeroplan Number: {str(e)}", "error")
                return None
                
        except Exception as e:
            self.log(f"✗ Error in form filling: {str(e)}", "error")
            return None
    
    def run(self):
        """Main thread execution"""
        try:
            self.log("=" * 60, "info")
            self.log("Starting CSV processing...", "info")
            self.log("=" * 60, "info")
            
            # Read CSV data
            self.log(f"Reading CSV file: {self.csv_file}", "info")
            data_list = self.read_csv(self.csv_file)
            self.log(f"✓ Found {len(data_list)} entries in CSV", "success")
            
            # Get already processed emails
            processed_emails = self.get_processed_emails()
            self.log(f"✓ Found {len(processed_emails)} already processed entries", "success")
            
            # Filter out already processed entries
            remaining_data = [row for row in data_list if row.get('Email') not in processed_emails]
            self.log(f"📋 {len(remaining_data)} entries remaining to process\n", "info")
            
            if not remaining_data:
                self.log("✨ All entries have been processed!", "success")
                self.finished_signal.emit()
                return
            
            # Process each row
            for index, data in enumerate(remaining_data, 1):
                if not self.is_running:
                    self.log("\n⚠️ Processing stopped by user", "warning")
                    break
                    
                email = data.get('Email')
                self.log("\n" + "=" * 60, "info")
                self.log(f"🔄 Processing entry {index}/{len(remaining_data)}", "info")
                self.log(f"📧 Email: {email}", "info")
                self.log(f"👤 Name: {data.get('First name')} {data.get('Last name')}", "info")
                self.log("=" * 60, "info")
                
                # Update progress
                self.progress_signal.emit(index, len(remaining_data))
                
                # Get proxy from CSV and format it
                proxy_raw = data.get('Proxy', '').strip()
                if proxy_raw:
                    proxy = self.format_proxy(proxy_raw)
                    if proxy:
                        self.log(f"🌐 Using proxy from CSV: {proxy}", "info")
                    else:
                        self.log(f"⚠️ Invalid proxy format in CSV: {proxy_raw}", "warning")
                        self.log("Skipping this entry...", "warning")
                        continue
                else:
                    self.log(f"⚠️ No proxy found in CSV for {email}", "warning")
                    self.log("Skipping this entry...", "warning")
                    continue
                
                try:
                    # Initialize browser with proxy
                    aeroplan_number = None
                    
                    # Check if extension directory exists
                    browser_args = {"uc": True, "proxy": proxy}
                    if os.path.exists(self.extension_dir):
                        browser_args["extension_dir"] = self.extension_dir
                    
                    # Check for custom Chrome binary
                    #chrome_binary = "chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing"
                    chrome_binary = "chrome-win64\\chrome.exe"
                    if os.path.exists(chrome_binary):
                        browser_args["binary_location"] = chrome_binary
                    
                    with SB(**browser_args) as sb:
                        aeroplan_number = self.fill_form(sb, data)
                    
                    # Save to processed CSV
                    self.save_processed_entry(data, aeroplan_number)
                    
                    if aeroplan_number:
                        self.log(f"✅ Successfully processed: {email}", "success")
                        self.log(f"✈️ Aeroplan Number: {aeroplan_number}", "success")
                    else:
                        self.log(f"⚠️ Processed but no Aeroplan number: {email}", "warning")
                    
                    # Remove from original CSV
                    self.remove_processed_row_from_csv(email)
                    
                except Exception as e:
                    self.log(f"❌ Error processing {email}: {str(e)}", "error")
                    self.log("⏭️ Skipping to next entry...", "warning")
                    continue
            
            self.log("\n" + "=" * 60, "info")
            self.log("🎉 All entries have been processed!", "success")
            self.log("=" * 60, "info")
            
        except Exception as e:
            self.log(f"❌ Critical error: {str(e)}", "error")
        finally:
            self.finished_signal.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.csv_file = None
        self.processed_file = "processed_data.csv"
        
        self.init_ui()
        
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Air Canada Aeroplan Registration Bot")
        self.setGeometry(100, 100, 1000, 700)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Title
        title = QLabel("🛫 Air Canada Aeroplan Registration Bot")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title)
        
        # Configuration Group
        config_group = QGroupBox("Configuration")
        config_layout = QVBoxLayout()
        
        # CSV File Selection
        csv_layout = QHBoxLayout()
        self.csv_label = QLabel("No CSV file selected (must include Proxy column)")
        self.csv_label.setStyleSheet("color: gray; padding: 5px;")
        csv_browse_btn = QPushButton("📁 Browse CSV File")
        csv_browse_btn.clicked.connect(self.browse_csv_file)
        csv_layout.addWidget(self.csv_label, 3)
        csv_layout.addWidget(csv_browse_btn, 1)
        config_layout.addLayout(csv_layout)
        
        # Info label about proxy
        proxy_info = QLabel("ℹ️ Proxies will be read from the 'Proxy' column in your CSV file")
        proxy_info.setStyleSheet("color: #FFC107; padding: 5px; font-style: italic;")
        config_layout.addWidget(proxy_info)
        
        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)
        
        # Control Buttons
        button_layout = QHBoxLayout()
        self.start_btn = QPushButton("▶️ Start Processing")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_processing)
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px; font-size: 14px;")
        
        self.stop_btn = QPushButton("⏸️ Stop Processing")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setStyleSheet("background-color: #f44336; color: white; padding: 10px; font-size: 14px;")
        
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.stop_btn)
        main_layout.addLayout(button_layout)
        
        # Progress Bar
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout()
        
        self.progress_label = QLabel("Ready to start")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        progress_group.setLayout(progress_layout)
        main_layout.addWidget(progress_group)
        
        # Log Output
        log_group = QGroupBox("Logs")
        log_layout = QVBoxLayout()
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: 'Courier New'; font-size: 12px;")
        log_layout.addWidget(self.log_output)
        
        clear_log_btn = QPushButton("🗑️ Clear Logs")
        clear_log_btn.clicked.connect(self.clear_logs)
        log_layout.addWidget(clear_log_btn)
        
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        # Status Bar
        self.statusBar().showMessage("Ready")
        
        # Log initial message
        self.add_log("Welcome to Air Canada Aeroplan Registration Bot!", "success")
        self.add_log("Please select a CSV file to begin.", "info")
    
    def browse_csv_file(self):
        """Open file dialog to select CSV file"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select CSV File",
            "",
            "CSV Files (*.csv);;All Files (*)"
        )
        
        if file_name:
            self.csv_file = file_name
            self.csv_label.setText(f"📄 {os.path.basename(file_name)}")
            self.csv_label.setStyleSheet("color: green; padding: 5px; font-weight: bold;")
            self.start_btn.setEnabled(True)
            self.add_log(f"CSV file selected: {file_name}", "success")
            
            # Count entries and check for Proxy column
            try:
                with open(file_name, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    fieldnames = reader.fieldnames
                    
                    # Check if Proxy column exists
                    if 'Proxy' not in fieldnames:
                        self.add_log("⚠️ Warning: 'Proxy' column not found in CSV!", "warning")
                        self.add_log("Please make sure your CSV has a 'Proxy' column with proxy data", "warning")
                        self.start_btn.setEnabled(False)
                        return
                    
                    count = sum(1 for row in reader if row.get('Email', '').strip())
                    self.add_log(f"Found {count} entries in the CSV file", "info")
                    self.add_log("✓ 'Proxy' column detected in CSV", "success")
            except Exception as e:
                self.add_log(f"Error reading CSV: {str(e)}", "error")
    
    def start_processing(self):
        """Start the CSV processing"""
        if not self.csv_file:
            self.add_log("Please select a CSV file first!", "error")
            return
        
        # Disable start button, enable stop button
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.csv_label.setEnabled(False)
        
        self.add_log("\n" + "=" * 60, "info")
        self.add_log("Starting processing...", "info")
        self.add_log("Proxies will be read from CSV 'Proxy' column", "info")
        self.add_log("=" * 60 + "\n", "info")
        
        # Create and start worker thread
        self.worker = WorkerThread(self.csv_file, self.processed_file)
        self.worker.log_signal.connect(self.add_log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.processing_finished)
        self.worker.start()
        
        self.statusBar().showMessage("Processing...")
    
    def stop_processing(self):
        """Stop the CSV processing"""
        if self.worker:
            self.add_log("\n⚠️ Stopping processing...", "warning")
            self.worker.stop()
            self.stop_btn.setEnabled(False)
    
    def processing_finished(self):
        """Called when processing is complete"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.csv_label.setEnabled(True)
        self.statusBar().showMessage("Ready")
        self.add_log("\n✅ Processing completed!", "success")
    
    def update_progress(self, current, total):
        """Update progress bar"""
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
        self.progress_label.setText(f"Processing {current} of {total} entries ({progress}%)")
    
    def add_log(self, message, level="info"):
        """Add a log message with color coding"""
        colors = {
            "info": "#d4d4d4",
            "success": "#4CAF50",
            "error": "#f44336",
            "warning": "#FFC107"
        }
        
        color = colors.get(level, colors["info"])
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        formatted_message = f'<span style="color: {color};">[{timestamp}] {message}</span>'
        self.log_output.append(formatted_message)
        
        # Auto-scroll to bottom
        cursor = self.log_output.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.log_output.setTextCursor(cursor)
    
    def clear_logs(self):
        """Clear the log output"""
        self.log_output.clear()
        self.add_log("Logs cleared.", "info")


def main():
    """Main function to run the GUI application"""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
