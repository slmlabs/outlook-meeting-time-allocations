from PyQt5.QtWidgets import QTableWidgetItem, QComboBox, QPushButton, QToolTip, QHeaderView
from PyQt5.QtCore import Qt
from outlook_handler import get_outlook_meetings
import pandas as pd


class MeetingManager:
    def __init__(self, meeting_table, notification_label, project_manager):
        self.meeting_table = meeting_table
        self.notification_label = notification_label
        self.project_manager = project_manager
        self.df = pd.DataFrame()
        self.dropdowns = []
        self.removed_meetings = set()
        self.setup_meeting_table()

    def setup_meeting_table(self):
        self.meeting_table.setColumnCount(4)
        self.meeting_table.setHorizontalHeaderLabels(
            ['Date & Time', 'Subject', 'Project', ''])
        self.meeting_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents)
        self.meeting_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.meeting_table.horizontalHeader().setSectionResizeMode(2,
                                                                   QHeaderView.Interactive)
        self.meeting_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
        # Set reasonable width for project column
        self.meeting_table.setColumnWidth(2, 200)
        # Set narrow width for remove column
        self.meeting_table.setColumnWidth(3, 30)

    def fetch_meetings(self, start_date, end_date):
        try:
            self.df = get_outlook_meetings(start_date, end_date)
            self.df = self.df.sort_values(by='Start')
            self.create_meeting_list()
            self.notification_label.setText(
                f"Fetched {len(self.df)} meetings.")
        except Exception as e:
            self.notification_label.setText(f"Error fetching meetings: {e}")

    def create_meeting_list(self):
        self.meeting_table.setRowCount(0)
        self.dropdowns = []  # Reset the dropdowns list
        self.removed_meetings = set()  # Reset the removed meetings set

        if not self.df.empty:
            for index, row in self.df.iterrows():
                self.meeting_table.insertRow(index)

                # Date & Time
                date_item = QTableWidgetItem(
                    row['Start'].strftime('%m/%d/%Y %H:%M'))
                self.meeting_table.setItem(index, 0, date_item)

                # Subject
                subject = row['Subject']
                truncated_subject = (
                    subject[:47] + '...') if len(subject) > 50 else subject
                subject_item = QTableWidgetItem(truncated_subject)
                subject_item.setToolTip(subject)
                self.meeting_table.setItem(index, 1, subject_item)

                # Project Selector
                combo = QComboBox()
                combo.addItems(self.project_manager.projects)
                combo.setCurrentText(row['Project'])  # Set to current project
                combo.currentIndexChanged.connect(
                    lambda index, subject=row['Subject'], combo=combo: self.update_project_assignment(subject, combo.currentText()))
                self.meeting_table.setCellWidget(index, 2, combo)
                self.dropdowns.append((row['Subject'], combo))

                # Delete Button
                remove_button = QPushButton("X")
                remove_button.setMaximumWidth(20)
                remove_button.setStyleSheet("text-align: center;")
                remove_button.clicked.connect(
                    lambda checked, row=index: self.remove_meeting(row))
                self.meeting_table.setCellWidget(index, 3, remove_button)

        else:
            self.notification_label.setText("No meetings found")

    def update_project_assignment(self, subject, project):
        self.df.loc[self.df['Subject'] == subject, 'Project'] = project
        self.notification_label.setStyleSheet("")  # Reset to default style
        self.notification_label.setText("Project assignment updated.")
        self.refresh_meeting_table()

    def remove_meeting(self, row):
        subject = self.df.iloc[row]['Subject']
        self.removed_meetings.add(subject)
        self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
        self.create_meeting_list()
        self.notification_label.setStyleSheet("")  # Reset to default style
        self.notification_label.setText(
            f"Meeting '{subject}' has been removed.")

    def refresh_meeting_table(self):
        for row in range(self.meeting_table.rowCount()):
            combo = self.meeting_table.cellWidget(row, 2)
            if combo is not None and isinstance(combo, QComboBox):
                project = self.df.iloc[row]['Project']
                combo.setCurrentText(project)

    def refresh_dropdowns(self):
        for subject, combo in self.dropdowns:
            current_value = combo.currentText()
            combo.clear()
            combo.addItems(self.project_manager.projects)
            combo.setCurrentText(
                current_value if current_value in self.project_manager.projects else "Unassigned")

    def check_unassigned(self):
        unassigned_count = self.df[self.df['Project'] == 'Unassigned'].shape[0]
        if unassigned_count > 0:
            self.notification_label.setStyleSheet(
                "color: red; font-size: 14px; font-weight: bold;")
            self.notification_label.setText(
                f"Warning: You have {unassigned_count} meetings with project 'Unassigned'.")
        else:
            self.notification_label.setStyleSheet("")
            self.notification_label.setText("")
