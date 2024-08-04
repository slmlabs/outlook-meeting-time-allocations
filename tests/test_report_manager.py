import unittest
from unittest.mock import MagicMock
from ui.report_manager import ReportManager
from PyQt5.QtWidgets import QApplication, QTextEdit
import pandas as pd
from datetime import datetime
import sys


class TestReportManager(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication(sys.argv)

    def setUp(self):
        self.meeting_manager = MagicMock()
        self.report_text_edit = QTextEdit()
        self.report_manager = ReportManager(
            self.meeting_manager, self.report_text_edit)

    def test_update_report(self):
        self.meeting_manager.df = pd.DataFrame({
            'Subject': ['Meeting 1', 'Meeting 2'],
            'Start': [datetime(2023, 8, 1, 9, 0), datetime(2023, 8, 1, 10, 0)],
            'End': [datetime(2023, 8, 1, 10, 0), datetime(2023, 8, 1, 11, 0)],
            'Duration': [1, 1],
            'Project': ['Unassigned', 'Project A']
        })
        self.meeting_manager.removed_meetings = set()
        self.report_manager.update_report()
        self.assertIn('Unassigned', self.report_text_edit.toHtml())
        self.assertIn('Project A', self.report_text_edit.toHtml())


if __name__ == '__main__':
    unittest.main()
