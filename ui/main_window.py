from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QVBoxLayout, QLabel, QHBoxLayout, QPushButton, QCalendarWidget, QTableWidget, QHeaderView, QTextEdit

from ui.meeting_manager import MeetingManager
from ui.project_manager import ProjectManager
from ui.report_manager import ReportManager

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meeting Assignment")
        self.centralWidget = QtWidgets.QWidget()
        self.setCentralWidget(self.centralWidget)
        self.layout = QVBoxLayout(self.centralWidget)

        self.setup_ui()

    def setup_ui(self):
        date_layout = QHBoxLayout()

        start_date_label = QLabel("Start Date:")
        self.start_date_entry = QCalendarWidget()
        date_layout.addWidget(start_date_label)
        date_layout.addWidget(self.start_date_entry)

        end_date_label = QLabel("End Date:")
        self.end_date_entry = QCalendarWidget()
        date_layout.addWidget(end_date_label)
        date_layout.addWidget(self.end_date_entry)

        self.layout.addLayout(date_layout)

        get_meetings_button = QPushButton("Get Meetings")
        get_meetings_button.clicked.connect(self.get_meetings)
        self.layout.addWidget(get_meetings_button)

        self.meeting_table = QTableWidget()
        self.meeting_table.setColumnCount(4)
        self.meeting_table.setHorizontalHeaderLabels(['Date & Time', 'Subject', 'Project', 'Remove'])
        self.meeting_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.meeting_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Date & Time column
        self.meeting_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Remove column
        self.layout.addWidget(self.meeting_table)

        self.notification_label = QLabel("")
        self.layout.addWidget(self.notification_label)

        # Initialize ProjectManager without meeting_manager reference
        self.project_manager = ProjectManager(self)
        self.layout.addLayout(self.project_manager.layout)

        # Initialize MeetingManager with reference to ProjectManager
        self.meeting_manager = MeetingManager(self.meeting_table, self.notification_label, self.project_manager)
        
        # Assign meeting_manager to project_manager after both are initialized
        self.project_manager.set_meeting_manager(self.meeting_manager)

        self.report_text_edit = QTextEdit()
        self.report_text_edit.setReadOnly(True)
        self.layout.addWidget(self.report_text_edit)

        update_report_button = QPushButton("Update Report")
        update_report_button.clicked.connect(self.update_report)
        self.layout.addWidget(update_report_button)

        self.report_manager = ReportManager(self.meeting_manager, self.report_text_edit)

    def get_meetings(self):
        try:
            start_date = self.start_date_entry.selectedDate().toPyDate()
            end_date = self.end_date_entry.selectedDate().toPyDate()
            self.meeting_manager.fetch_meetings(start_date, end_date)
        except Exception as e:
            self.notification_label.setText(f"Error fetching meetings: {e}")

    def update_report(self):
        self.report_manager.update_report()
