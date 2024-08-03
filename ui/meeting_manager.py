from PyQt5.QtWidgets import QTableWidgetItem, QComboBox, QPushButton, QLabel, QWidget, QHBoxLayout
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

    def fetch_meetings(self, start_date, end_date):
        self.df = get_outlook_meetings(start_date, end_date)
        self.df = self.df.sort_values(by='Start')
        self.create_meeting_list()

    def create_meeting_list(self):
        self.meeting_table.setRowCount(0)
        self.dropdowns = []  # Reset the dropdowns list
        self.removed_meetings = set()  # Reset the removed meetings set

        if not self.df.empty:
            for index, row in self.df.iterrows():
                self.meeting_table.insertRow(index)

                # Date & Time
                date_item = QTableWidgetItem(row['Start'].strftime('%m/%d/%Y %H:%M'))
                self.meeting_table.setItem(index, 0, date_item)

                # Subject
                subject_item = QTableWidgetItem(row['Subject'])
                self.meeting_table.setItem(index, 1, subject_item)

                # Project Selector and Badge
                container_widget = QWidget()
                container_layout = QHBoxLayout(container_widget)
                combo = QComboBox()
                combo.addItems(self.project_manager.projects)
                combo.setCurrentText("Unassigned")  # Default to "Unassigned"
                combo.currentIndexChanged.connect(lambda index, subject=row['Subject'], combo=combo: self.update_project_assignment(subject, combo.currentText()))
                container_layout.addWidget(combo)

                badge = QLabel()
                badge.setStyleSheet("background-color: red; color: white; border-radius: 4px; padding: 2px;")
                badge.setText("!")
                badge.setVisible(combo.currentText() == "Unassigned")
                container_layout.addWidget(badge)

                container_layout.setContentsMargins(0, 0, 0, 0)
                self.meeting_table.setCellWidget(index, 2, container_widget)
                self.dropdowns.append((row['Subject'], combo, badge))

                # Delete Button
                remove_button = QPushButton("X")
                remove_button.clicked.connect(lambda checked, row=index: self.remove_meeting(row))
                self.meeting_table.setCellWidget(index, 3, remove_button)

        else:
            self.notification_label.setText("No meetings found")

    def update_project_assignment(self, subject, project):
        self.df.loc[self.df['Subject'] == subject, 'Project'] = project
        self.notification_label.setText("Project assignment updated.")
        for subject_, combo, badge in self.dropdowns:
            if subject_ == subject:
                badge.setVisible(project == "Unassigned")

    def remove_meeting(self, row):
        subject = self.df.iloc[row]['Subject']
        self.removed_meetings.add(subject)
        self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
        self.create_meeting_list()
        self.notification_label.setText(f"Meeting '{subject}' has been removed.")

    def refresh_dropdowns(self):
        for subject, combo, badge in self.dropdowns:
            current_value = combo.currentText()
            combo.clear()
            combo.addItems(self.project_manager.projects)
            combo.setCurrentText(current_value if current_value in self.project_manager.projects else "Unassigned")
            badge.setVisible(combo.currentText() == "Unassigned")

    def check_unassigned(self):
        unassigned_count = self.df[self.df['Project'] == 'Unassigned'].shape[0]
        if unassigned_count > 0:
            self.notification_label.setStyleSheet("color: red; font-size: 14px; font-weight: bold;")
            self.notification_label.setText(f"Warning: You have {unassigned_count} meetings with project 'Unassigned'.")
        else:
            self.notification_label.setStyleSheet("")
            self.notification_label.setText("")
