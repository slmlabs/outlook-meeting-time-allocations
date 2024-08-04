import unittest
from unittest.mock import MagicMock, patch
from PyQt5.QtWidgets import QApplication, QTableWidget, QLabel
from ui.meeting_manager import MeetingManager
import pandas as pd
from datetime import datetime
import sys


class TestMeetingManager(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication(sys.argv)

    def setUp(self):
        self.meeting_table = QTableWidget()
        self.notification_label = QLabel()
        self.project_manager = MagicMock()
        self.project_manager.projects = [
            'Project A', 'Project B', 'Unassigned']
        self.meeting_manager = MeetingManager(
            self.meeting_table, self.notification_label, self.project_manager)

    @patch('ui.meeting_manager.get_outlook_meetings')
    def test_fetch_meetings(self, mock_get_outlook_meetings):
        mock_get_outlook_meetings.return_value = pd.DataFrame({
            'Subject': ['Meeting 1', 'Meeting 2'],
            'Start': [datetime(2023, 8, 1, 9, 0), datetime(2023, 8, 1, 10, 0)],
            'End': [datetime(2023, 8, 1, 10, 0), datetime(2023, 8, 1, 11, 0)],
            'Duration': [1, 1],
            'Project': ['Unassigned', 'Unassigned']
        })
        self.meeting_manager.fetch_meetings(
            datetime(2023, 8, 1), datetime(2023, 8, 7))
        self.assertEqual(self.meeting_manager.df.shape[0], 2)
        self.assertEqual(
            self.meeting_manager.df.iloc[0]['Subject'], 'Meeting 1')

    def test_update_project_assignment(self):
        self.meeting_manager.df = pd.DataFrame({
            'Subject': ['Meeting 1'],
            'Project': ['Unassigned']
        })
        self.meeting_manager.update_project_assignment(
            'Meeting 1', 'Project A')
        self.assertEqual(
            self.meeting_manager.df.iloc[0]['Project'], 'Project A')

    def test_check_unassigned(self):
        self.meeting_manager.df = pd.DataFrame({
            'Subject': ['Meeting 1', 'Meeting 2'],
            'Project': ['Unassigned', 'Project A']
        })
        self.meeting_manager.check_unassigned()
        self.assertIn('Warning', self.notification_label.text())


if __name__ == '__main__':
    unittest.main()
