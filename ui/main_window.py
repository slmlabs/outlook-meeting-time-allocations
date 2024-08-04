from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QVBoxLayout, QLabel, QHBoxLayout, QPushButton, QDateEdit, QTableWidget, QHeaderView, QTextEdit, QTableWidgetItem
import psutil
import subprocess
import os
import json

from ui.meeting_manager import MeetingManager
from ui.project_manager import ProjectManager
from ui.report_manager import ReportManager

PROJECTS_FILE = 'projects.json'


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meeting Assignment")
        self.centralWidget = QtWidgets.QWidget()
        self.setCentralWidget(self.centralWidget)
        self.layout = QVBoxLayout(self.centralWidget)

        self.setup_ui()
        self.setup_timer()
        self.check_outlook_status()

    def setup_ui(self):
        # Outlook status layout
        self.outlook_status_layout = QHBoxLayout()
        self.outlook_status_label = QLabel("")
        self.outlook_status_layout.addWidget(self.outlook_status_label)
        self.outlook_button = QPushButton("Open Outlook")
        self.outlook_button.clicked.connect(self.open_outlook)
        self.outlook_status_layout.addWidget(self.outlook_button)
        self.layout.addLayout(self.outlook_status_layout)

        # Date range selection
        date_range_layout = QHBoxLayout()

        self.start_date_edit = QDateEdit(QtCore.QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("MM/dd/yyyy")
        date_range_layout.addWidget(QLabel("Start Date:"))
        date_range_layout.addWidget(self.start_date_edit)

        self.end_date_edit = QDateEdit(QtCore.QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("MM/dd/yyyy")
        date_range_layout.addWidget(QLabel("End Date:"))
        date_range_layout.addWidget(self.end_date_edit)

        self.layout.addLayout(date_range_layout)

        get_meetings_button = QPushButton("Get Meetings")
        get_meetings_button.clicked.connect(self.get_meetings)
        self.layout.addWidget(get_meetings_button)

        self.meeting_table = QTableWidget()
        self.meeting_table.setColumnCount(5)
        self.meeting_table.setHorizontalHeaderLabels(
            ['Unassigned', 'Date & Time', 'Subject', 'Project', 'Remove'])
        self.meeting_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.meeting_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents)  # Unassigned column
        self.meeting_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeToContents)  # Date & Time column
        self.meeting_table.horizontalHeader().setSectionResizeMode(
            4, QHeaderView.ResizeToContents)  # Remove column
        self.layout.addWidget(self.meeting_table)

        self.notification_label = QLabel("")
        self.layout.addWidget(self.notification_label)

        # Initialize ProjectManager without meeting_manager reference
        self.project_manager = ProjectManager(self)
        self.layout.addWidget(self.project_manager)

        # Initialize MeetingManager with reference to ProjectManager
        self.meeting_manager = MeetingManager(
            self.meeting_table, self.notification_label, self.project_manager)

        # Assign meeting_manager to project_manager after both are initialized
        self.project_manager.set_meeting_manager(self.meeting_manager)

        self.report_text_edit = QTextEdit()
        self.report_text_edit.setReadOnly(True)
        self.layout.addWidget(self.report_text_edit)

        update_report_button = QPushButton("Update Report")
        update_report_button.clicked.connect(self.update_report)
        self.layout.addWidget(update_report_button)

        self.report_manager = ReportManager(
            self.meeting_manager, self.report_text_edit)

    def setup_timer(self):
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.check_outlook_status)
        self.timer.start(5000)  # Check Outlook status every 5 seconds

    def get_meetings(self):
        try:
            start_date = self.start_date_edit.date().toPyDate()
            end_date = self.end_date_edit.date().toPyDate()
            if start_date > end_date:
                self.notification_label.setText(
                    "Start date must be before end date.")
                return
            self.meeting_manager.fetch_meetings(start_date, end_date)
        except Exception as e:
            self.notification_label.setText(f"Error fetching meetings: {e}")

    def update_report(self):
        self.report_manager.update_report()

    def check_outlook_status(self):
        if self.is_outlook_running():
            self.outlook_status_label.setText("Ready")
            self.outlook_status_label.setStyleSheet(
                "color: green; font-weight: bold;")
            self.outlook_button.hide()
        else:
            self.outlook_status_label.setText("Outlook is not running.")
            self.outlook_status_label.setStyleSheet(
                "color: red; font-weight: bold;")
            self.outlook_button.show()

    def is_outlook_running(self):
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'OUTLOOK.EXE':
                return True
        return False

    def open_outlook(self):
        outlook_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",  # Office 365
            # Office 2016
            r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
            # Office 2013
            r"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE",
            # Office 2010
            r"C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE",
        ]
        for path in outlook_paths:
            if os.path.exists(path):
                try:
                    subprocess.Popen([path])
                    # Check Outlook status after a short delay to allow it to start
                    QtCore.QTimer.singleShot(5000, self.check_outlook_status)
                    return
                except Exception as e:
                    self.notification_label.setText(
                        f"Error opening Outlook: {e}")
                    return
        self.notification_label.setText("Outlook executable not found.")

    def refresh_meeting_table(self):
        self.meeting_table.setRowCount(len(self.meeting_manager.df))
        for row, meeting in self.meeting_manager.df.iterrows():
            unassigned_dot = QLabel()
            if meeting['Project'] == 'Unassigned':
                unassigned_dot.setPixmap(QtGui.QPixmap(10, 10))
                unassigned_dot.setStyleSheet(
                    "background-color: red; border-radius: 5px;")
            self.meeting_table.setCellWidget(row, 0, unassigned_dot)
            self.meeting_table.setItem(row, 1, QTableWidgetItem(
                meeting['Start'].strftime("%Y-%m-%d %H:%M")))
            self.meeting_table.setItem(
                row, 2, QTableWidgetItem(meeting['Subject']))
            project_dropdown = self.create_project_dropdown(meeting['Project'])
            self.meeting_table.setCellWidget(row, 3, project_dropdown)
            remove_button = QPushButton("X")
            remove_button.clicked.connect(
                lambda _, r=row: self.remove_meeting(r))
            self.meeting_table.setCellWidget(row, 4, remove_button)

    def create_project_dropdown(self, selected_project):
        dropdown = QtWidgets.QComboBox()
        dropdown.addItems(self.project_manager.projects)
        dropdown.setCurrentText(selected_project)
        dropdown.currentIndexChanged.connect(
            lambda: self.update_meeting_project(dropdown))
        return dropdown

    def update_meeting_project(self, dropdown):
        row = self.meeting_table.indexAt(dropdown.pos()).row()
        project = dropdown.currentText()
        self.meeting_manager.df.at[row, 'Project'] = project
        self.refresh_meeting_table()

    def remove_meeting(self, row):
        self.meeting_manager.df = self.meeting_manager.df.drop(
            row).reset_index(drop=True)
        self.refresh_meeting_table()


if __name__ == '__main__':
    from PyQt5.QtWidgets import QApplication
    import sys

    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
