import sys
import subprocess
import threading
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton,QMessageBox, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QMenu, QAction, QDialog, QLabel, QFileDialog
from PyQt5.QtCore import Qt, QTranslator, QLocale, QUrl
from PyQt5.QtGui import QDesktopServices
import pandas as pd

class WifiProfileViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.profiles = []
        self.current_profile_index = 0
        self.thread = None
        self.show()

    def initUI(self):
        self.setWindowTitle('Wi-Fi Profile Viewer v1.0')
        self.setGeometry(100, 100, 600, 400)

        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['SSID', 'Password'])

        self.refresh_button = QPushButton('Refresh', self)
        self.refresh_button.clicked.connect(self.refresh_profiles)

        self.save_button = QPushButton('Save to Excel', self)
        self.save_button.clicked.connect(self.save_to_excel)

        self.total_data_label = QLabel(self)

        layout = QVBoxLayout()
        layout.addWidget(self.refresh_button)
        layout.addWidget(self.save_button)
        layout.addWidget(self.table)
        layout.addWidget(self.total_data_label)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Create a Help menu
        help_menu = self.menuBar().addMenu('Help')
        visit_website_action = QAction('Visit Website', self)
        visit_website_action.triggered.connect(self.visit_website)
        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(visit_website_action)
        help_menu.addAction(about_action)

    def refresh_profiles(self):
        if self.thread is None or not self.thread.is_alive():
            self.current_profile_index = 0
            self.profiles = []
            self.table.setRowCount(0)
            self.thread = threading.Thread(target=self.get_profiles)
            self.thread.start()

    def get_profiles(self):
        data = subprocess.check_output(['netsh', 'wlan', 'show', 'profiles']).decode('utf-8', errors="backslashreplace").split('\n')
        profiles = [i.split(":")[1][1:-1] for i in data if "All User Profile" in i]

        for profile_name in profiles:
            try:
                results = subprocess.check_output(['netsh', 'wlan', 'show', 'profile', profile_name, 'key=clear']).decode('utf-8', errors="backslashreplace").split('\n')
                results = [b.split(":")[1][1:-1] for b in results if "Key Content" in b]
                password = results[0] if results else ''
                self.profiles.append((profile_name, password))
                self.update_table()
            except subprocess.CalledProcessError:
                self.profiles.append((profile_name, 'ENCODING ERROR'))
                self.update_table()

    def update_table(self):
        if self.current_profile_index < len(self.profiles):
            profile_name, password = self.profiles[self.current_profile_index]
            self.table.insertRow(self.current_profile_index)
            self.table.setItem(self.current_profile_index, 0, QTableWidgetItem(profile_name))
            self.table.setItem(self.current_profile_index, 1, QTableWidgetItem(password))
            self.current_profile_index += 1
            self.update_total_data_label()

    def update_total_data_label(self):
        total_profiles = len(self.profiles)
        self.total_data_label.setText(f'Total Data: {total_profiles}')

    def visit_website(self):
        QDesktopServices.openUrl(QUrl("http://www.sihabsahariar.com"))

    def show_about_dialog(self):
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle('About')
        about_dialog.setWindowFlag(Qt.WindowStaysOnTopHint)

        about_label = QLabel('Wi-Fi Profile Viewer\n\nCreated by: Sihab Sahariar\nVersion: 1.0')
        ok_button = QPushButton('OK')
        ok_button.clicked.connect(about_dialog.accept)

        layout = QVBoxLayout()
        layout.addWidget(about_label)
        layout.addWidget(ok_button)
        about_dialog.setLayout(layout)

        about_dialog.exec_()

    def save_to_excel(self):
        if self.profiles:
            data = {'SSID': [], 'Password': []}
            for profile_name, password in self.profiles:
                data['SSID'].append(profile_name)
                data['Password'].append(password)

            df = pd.DataFrame(data)

            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            selected_file, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)

            if selected_file:
                if not selected_file.endswith('.xlsx'):
                    selected_file += '.xlsx'  # Ensure the file has the .xlsx extension

                try:
                    df.to_excel(selected_file, index=False)
                    QMessageBox.information(self, "Success", "File saved successfully.")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

def main():
    app = QApplication(sys.argv)

    # Set the dark style
    app.setStyle('Fusion')
    dark_palette = app.palette()
    dark_palette.setColor(dark_palette.Window, Qt.darkGray)
    dark_palette.setColor(dark_palette.WindowText, Qt.white)
    dark_palette.setColor(dark_palette.Button, Qt.darkGray)
    dark_palette.setColor(dark_palette.ButtonText, Qt.white)
    dark_palette.setColor(dark_palette.Window, Qt.darkGray)
    dark_palette.setColor(dark_palette.WindowText, Qt.white)
    dark_palette.setColor(dark_palette.Base, Qt.darkGray)
    dark_palette.setColor(dark_palette.AlternateBase, Qt.darkGray)
    dark_palette.setColor(dark_palette.ToolTipBase, Qt.white)
    dark_palette.setColor(dark_palette.ToolTipText, Qt.white)
    app.setPalette(dark_palette)

    window = WifiProfileViewer()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
