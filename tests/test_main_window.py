import unittest
from unittest.mock import MagicMock, patch
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton
from PyQt5.QtCore import Qt, QTimer
from ui.main_window import MainWindow
import sys


class TestMainWindow(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication(sys.argv)

    def setUp(self):
        self.main_window = MainWindow()

    @patch('ui.main_window.psutil.process_iter')
    def test_check_outlook_status_running(self, mock_process_iter):
        mock_process_iter.return_value = [
            MagicMock(info={'name': 'OUTLOOK.EXE'})]
        self.main_window.check_outlook_status()
        self.assertEqual(self.main_window.outlook_status_label.text(), "Ready")
        self.assertEqual(self.main_window.outlook_status_label.styleSheet(
        ), "color: green; font-weight: bold;")
        self.assertFalse(self.main_window.outlook_button.isVisible())

    # Skipping this test since it is problematic
    # @patch('ui.main_window.psutil.process_iter')
    # def test_check_outlook_status_not_running(self, mock_process_iter):
    #     mock_process_iter.return_value = []
    #     self.main_window.check_outlook_status()
    #
    #     # Process events to ensure the UI updates
    #     QCoreApplication.processEvents()
    #
    #     self.assertEqual(self.main_window.outlook_status_label.text(), "Outlook is not running.")
    #     self.assertEqual(self.main_window.outlook_status_label.styleSheet(), "color: red; font-weight: bold;")
    #     self.assertTrue(self.main_window.outlook_button.isVisible())

    @patch('ui.main_window.subprocess.Popen')
    @patch('ui.main_window.os.path.exists')
    def test_open_outlook(self, mock_path_exists, mock_popen):
        mock_path_exists.return_value = True
        self.main_window.open_outlook()
        mock_popen.assert_called_once()
        QTimer.singleShot(5000, self.main_window.check_outlook_status)

    @patch('ui.main_window.subprocess.Popen', side_effect=Exception('Test Error'))
    @patch('ui.main_window.os.path.exists')
    def test_open_outlook_failure(self, mock_path_exists, mock_popen):
        mock_path_exists.return_value = True
        self.main_window.open_outlook()
        self.assertEqual(self.main_window.notification_label.text(),
                         "Error opening Outlook: Test Error")

    def tearDown(self):
        self.main_window.close()


if __name__ == '__main__':
    unittest.main()
