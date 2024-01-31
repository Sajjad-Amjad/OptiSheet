import sys
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel,
                             QComboBox, QLineEdit, QProgressBar, QAction, QHBoxLayout, QCheckBox,
                             QDialog, QFileDialog, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal
import openai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import configparser
from PyQt5.QtCore import Qt
from io import StringIO
import requests


# Function to process data using OpenAI and update Google Sheet per row
def process_data_with_openai(text_content, analysis_instruction, openai_key):
    openai.api_key = openai_key

    try:
        prompt = f"{analysis_instruction}: '{text_content}'? (Answer with Yes or No)"
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=10
        )
        return response.choices[0].message['content'].strip()
    except Exception as e:
        return f"Error: {e}"

# Function to load data from Google Sheets


def load_data_from_sheet(sheet_url, credentials_path):
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        credentials_path, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(sheet_url).sheet1
    data = sheet.get_all_values()
    headers = data.pop(0)
    return pd.DataFrame(data, columns=headers), sheet

# Worker class for loading data


class DataLoader(QThread):
    finished = pyqtSignal(pd.DataFrame, object, str)

    def __init__(self, sheet_url, credentials_path):
        super().__init__()
        self.sheet_url = sheet_url
        self.credentials_path = credentials_path

    def run(self):
        try:
            df, sheet = load_data_from_sheet(
                self.sheet_url, self.credentials_path)
            self.finished.emit(df, sheet, "")
        except Exception as e:
            self.finished.emit(pd.DataFrame(), None, str(e))

# Worker class for processing data


class DataProcessor(QThread):
    update_progress = pyqtSignal(int)
    processing_complete = pyqtSignal(str)

    def __init__(self, sheet, df, text_col, instr_col_or_manual, result_col_index, openai_key, use_manual_instr=False):
        super().__init__()
        self.sheet = sheet
        self.df = df
        self.text_col = text_col
        self.instr_col_or_manual = instr_col_or_manual
        self.result_col_index = result_col_index
        self.openai_key = openai_key
        self.use_manual_instr = use_manual_instr

    def run(self):
        try:
            for index, row in self.df.iterrows():
                instruction = self.instr_col_or_manual if self.use_manual_instr else row[
                    self.instr_col_or_manual]
                result = process_data_with_openai(
                    row[self.text_col], instruction, self.openai_key)
                if "Error:" not in result:
                    if self.sheet:
                        self.sheet.update_cell(
                            index + 2, self.result_col_index, result)
                    else:
                        self.df.at[index,
                                   self.df.columns[self.result_col_index - 1]] = result
                self.update_progress.emit(index + 1)
            self.processing_complete.emit("")
        except Exception as e:
            self.processing_complete.emit(str(e))

# Configuration window


class ConfigWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Configuration Settings')
        self.setGeometry(100, 100, 400, 200)
        layout = QVBoxLayout()

        # Google Sheets Credentials File Path
        self.credentials_path = QLineEdit(self)
        self.credentials_button = QPushButton("Upload Credentials File", self)
        self.credentials_button.clicked.connect(self.upload_credentials)
        layout.addWidget(QLabel('Google Sheets Credentials File Path:'))
        layout.addWidget(self.credentials_path)
        layout.addWidget(self.credentials_button)

        # OpenAI API Key
        self.openai_api_key_input = QLineEdit(self)
        layout.addWidget(QLabel('OpenAI API Key:'))
        layout.addWidget(self.openai_api_key_input)

        # Save Button
        save_button = QPushButton('Save', self)
        save_button.clicked.connect(self.save_configs)
        layout.addWidget(save_button)

        self.setLayout(layout)

    def upload_credentials(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open Credentials File", "", "JSON Files (*.json)")
        if file_name:
            self.credentials_path.setText(file_name)

    def save_configs(self):
        configs = configparser.ConfigParser()
        configs['DEFAULT'] = {
            'CredentialsPath': self.credentials_path.text(),
            'OpenAIKey': self.openai_api_key_input.text()
        }
        with open('config.ini', 'w') as configfile:
            configs.write(configfile)
        self.close()

# Main application window


class SheetProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.save_file_location = None
        try:
            self.configs = self.load_configs()
            self.initUI()
        except Exception as e:
            self.showErrorDialog(f"Initialization Error: {e}")

    def load_configs(self):
        try:
            configs = configparser.ConfigParser()
            configs.read('config.ini')
            return configs
        except Exception as e:
            self.showErrorDialog(f"Config Load Error: {e}")
            raise e

    def showSaveDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self,
                                                  "Save Processed Data", "",
                                                  "CSV Files (*.csv);;All Files (*)", options=options)
        if fileName:
            if not fileName.endswith('.csv'):
                fileName += '.csv'
            return fileName
        return None

    def initUI(self):
        try:
            self.setWindowTitle('OptiSheet')
            self.setGeometry(100, 100, 800, 400)
            layout = QVBoxLayout()

            # Menu bar for configurations
            menubar = self.menuBar()
            fileMenu = menubar.addMenu('File')
            configAction = QAction('Configurations', self)
            configAction.triggered.connect(self.openConfigWindow)
            fileMenu.addAction(configAction)

            # Option to choose data source
            self.data_source_label = QLabel("Data Source:")
            self.data_source_combo = QComboBox(self)
            self.data_source_combo.addItems(
                ["Google Sheet", "Google Sheet CSV URL", "Local CSV"])
            self.data_source_combo.currentIndexChanged.connect(
                self.onDataSourceChanged)

            layout.addWidget(self.data_source_label)
            layout.addWidget(self.data_source_combo)

            # URL input
            self.url_input = QLineEdit(self)
            layout.addWidget(QLabel('Google Sheet URL:'))
            layout.addWidget(self.url_input)

            # Get Data button
            self.get_data_button = QPushButton('Get Data', self)
            self.get_data_button.clicked.connect(self.loadData)
            layout.addWidget(self.get_data_button)

            # Dropdowns for column selection
            layout.addWidget(QLabel('Text Column:'))
            self.text_col_dropdown = QComboBox(self)
            layout.addWidget(self.text_col_dropdown)
            layout.addWidget(QLabel('Instruction Column:'))
            self.instr_col_dropdown = QComboBox(self)
            layout.addWidget(self.instr_col_dropdown)

            # Manual instruction input
            self.manual_instr_checkbox = QCheckBox("Use manual instruction")
            self.manual_instr_input = QLineEdit(self)
            self.manual_instr_input.setPlaceholderText(
                "Enter instruction here")
            layout.addWidget(self.manual_instr_checkbox)
            layout.addWidget(self.manual_instr_input)

            layout.addWidget(QLabel('Result Column:'))
            self.result_col_dropdown = QComboBox(self)
            layout.addWidget(self.result_col_dropdown)

            # Button for save file location
            self.save_location_button = QPushButton(
                'Select Save Location', self)
            self.save_location_button.clicked.connect(self.selectSaveLocation)
            self.save_location_button.setDisabled(True)  # Disabled by default
            layout.addWidget(self.save_location_button)

            # Process button
            self.process_button = QPushButton('Process', self)
            self.process_button.clicked.connect(self.processData)
            layout.addWidget(self.process_button)

            # Progress bar
            self.progress_bar = QProgressBar(self)
            layout.addWidget(self.progress_bar)

            # Status label
            self.status_label = QLabel('Status: Ready', self)
            layout.addWidget(self.status_label)

            # Set main widget
            main_widget = QWidget()
            main_widget.setLayout(layout)
            self.setCentralWidget(main_widget)
        except Exception as e:
            self.showErrorDialog(f"UI Initialization Error: {e}")

    def selectSaveLocation(self):
        self.save_file_location = self.showSaveDialog()
        if self.save_file_location:
            self.status_label.setText(
                f'Save location selected: {self.save_file_location}')
        else:
            self.status_label.setText('Save location not selected.')

    def openConfigWindow(self):
        self.config_window = ConfigWindow()
        self.config_window.show()

    def loadGoogleSheet(self):
        try:
            self.progress_bar.setRange(0, 0)
            credentials_path = self.configs['DEFAULT'].get(
                'CredentialsPath', '')
            self.data_loader = DataLoader(
                self.url_input.text(), credentials_path)
            self.data_loader.finished.connect(self.onDataLoaded)
            self.data_loader.start()
        except Exception as e:
            self.showErrorDialog(f"Load Data Error: {e}")

    def loadCSVFromURL(self, url):
        try:
            response = requests.get(url)
            response.raise_for_status()
            csv_data = response.content.decode('utf-8')
            self.df = pd.read_csv(StringIO(csv_data))
            self.sheet = None
            self.populateColumnDropdowns()
            self.progress_bar.setRange(0, 1)
            self.status_label.setText('CSV Loaded')
            self.save_location_button.setDisabled(False)
        except Exception as e:
            self.showErrorDialog(f"Error Loading CSV from URL: {e}")

    def loadData(self):
        data_source = self.data_source_combo.currentText()
        if data_source == "Local CSV":
            self.uploadCSV()
        elif data_source == "Google Sheet":
            self.loadGoogleSheet()
        else:  # Assuming "Google Sheet CSV URL"
            self.loadCSVFromURL(self.url_input.text())

    def uploadCSV(self):
        try:
            file_name, _ = QFileDialog.getOpenFileName(
                self, "Open CSV File", "", "CSV Files (*.csv)")
            if file_name:
                self.df = pd.read_csv(file_name)
                self.sheet = None
                self.populateColumnDropdowns()
                self.progress_bar.setRange(0, 1)
                self.status_label.setText('CSV Loaded')
                self.save_location_button.setDisabled(False)
        except Exception as e:
            self.showErrorDialog(f"CSV Upload Error: {e}")

    def populateColumnDropdowns(self):
        self.text_col_dropdown.clear()
        self.instr_col_dropdown.clear()
        self.result_col_dropdown.clear()

        for col in self.df.columns:
            self.text_col_dropdown.addItem(col)
            self.instr_col_dropdown.addItem(col)
            self.result_col_dropdown.addItem(col)

        if not self.sheet:  # If loading from CSV or CSV URL
            result_col_name = self.result_col_dropdown.currentText()
            if result_col_name in self.df.columns:
                self.df[result_col_name] = self.df[result_col_name].astype(
                    'object')

    def onDataLoaded(self, df, sheet, error):
        if error:
            self.showErrorDialog(f"Data Loading Error: {error}")
            self.progress_bar.setRange(0, 1)
        else:
            self.df = df
            self.sheet = sheet
            for col in df.columns:
                self.text_col_dropdown.addItem(col)
                self.instr_col_dropdown.addItem(col)
                self.result_col_dropdown.addItem(col)
            self.progress_bar.setRange(0, 1)
            self.status_label.setText('Data Loaded')

    def processData(self):
        try:
            if self.sheet:
                if hasattr(self, 'df') and hasattr(self, 'sheet'):
                    self.progress_bar.setMaximum(len(self.df))
                    openai_key = self.configs['DEFAULT'].get('OpenAIKey', '')
                    result_col_index = self.df.columns.get_loc(
                        self.result_col_dropdown.currentText()) + 1
                    if self.manual_instr_checkbox.isChecked():
                        manual_instruction = self.manual_instr_input.text().strip()
                        self.data_processor = DataProcessor(self.sheet, self.df, self.text_col_dropdown.currentText(),
                                                            manual_instruction, result_col_index, openai_key, use_manual_instr=True)
                    else:
                        self.data_processor = DataProcessor(self.sheet, self.df, self.text_col_dropdown.currentText(),
                                                            self.instr_col_dropdown.currentText(), result_col_index, openai_key)
                    self.data_processor.update_progress.connect(
                        self.onUpdateProgress)
                    self.data_processor.processing_complete.connect(
                        self.onDataProcessed)
                    self.data_processor.start()
            else:
                # Logic for processing CSV
                if not self.save_file_location:  # Check if save location is selected for CSV
                    QMessageBox.warning(
                        self, "Warning", "Please select a location to save the file first.")
                    return

                openai_key = self.configs['DEFAULT'].get('OpenAIKey', '')
                result_col_index = self.df.columns.get_loc(
                    self.result_col_dropdown.currentText()) + 1
                self.data_processor = DataProcessor(None, self.df, self.text_col_dropdown.currentText(),
                                                    self.instr_col_dropdown.currentText(), result_col_index, openai_key)
                self.data_processor.update_progress.connect(
                    self.onUpdateProgress)
                self.data_processor.processing_complete.connect(
                    self.onDataProcessed)
                self.data_processor.start()
        except Exception as e:
            self.showErrorDialog(f"Process Data Error: {e}")

    def onUpdateProgress(self, processed_rows):
        self.progress_bar.setValue(processed_rows)

    def onDataProcessed(self, error):
        if error:
            self.showErrorDialog(f"Data Processing Error: {error}")
        else:
            if not self.sheet:  # If processing CSV
                if self.save_file_location:  # Check if the save location was set
                    try:
                        # Save the updated DataFrame to the previously chosen file location
                        self.df.to_csv(self.save_file_location, index=False)
                        self.status_label.setText(
                            f'Processing Complete. Data saved to "{self.save_file_location}"')
                    except Exception as e:
                        self.showErrorDialog(f"Error saving CSV: {e}")
                else:
                    # This case should not normally occur as the location should be set before processing
                    self.showErrorDialog(
                        "No save location was selected for the CSV data.")
            else:
                self.status_label.setText('Processing Complete')

        self.progress_bar.setRange(0, 1)

    def onDataSourceChanged(self, index):
        data_source = self.data_source_combo.currentText()
        if data_source == "Local CSV":
            self.url_input.setDisabled(True)
            self.get_data_button.setText("Upload CSV")
        else:
            self.url_input.setDisabled(False)
            self.get_data_button.setText("Get Data")

    def showErrorDialog(self, message):
        QMessageBox.critical(self, "Error", message)


if __name__ == '__main__':
    QApplication.setAttribute(
        Qt.AA_DontUseNativeMenuBar, True)  # Add this line
    app = QApplication(sys.argv)
    mainWin = SheetProcessorApp()
    mainWin.show()
    sys.exit(app.exec_())
