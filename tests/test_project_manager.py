import unittest
from unittest.mock import MagicMock, mock_open, patch
from PyQt5.QtWidgets import QApplication, QLabel
from ui.project_manager import ProjectManager
import sys
import json
import os

TEST_PROJECT_FILE = 'test_projects.json'


class TestProjectManager(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication(sys.argv)

    def setUp(self):
        self.parent = MagicMock()
        self.parent.notification_label = QLabel()
        self.mock_projects = ['Unassigned', 'Project A', 'Project B']
        self.mock_file_data = json.dumps(self.mock_projects)
        self.mock_open = mock_open(read_data=self.mock_file_data)
        with patch('builtins.open', self.mock_open):
            self.project_manager = ProjectManager(
                self.parent, project_file=TEST_PROJECT_FILE)

    @patch('builtins.open', new_callable=mock_open, read_data='["Unassigned", "Project A", "Project B"]')
    def test_add_project(self, mock_file):
        with patch('json.dump'):
            result = self.project_manager.add_project_direct('Project C')
            self.assertTrue(result)
            self.assertIn('Project C', self.project_manager.projects)

    @patch('builtins.open', new_callable=mock_open, read_data='["Unassigned", "Project A", "Project B"]')
    def test_remove_project(self, mock_file):
        self.project_manager.projects.append('Project D')
        with patch('json.dump'):
            result = self.project_manager.remove_project_direct('Project D')
            self.assertTrue(result)
            self.assertNotIn('Project D', self.project_manager.projects)

    @patch('builtins.open', new_callable=mock_open, read_data='["Unassigned", "Project A", "Project B"]')
    def test_update_project(self, mock_file):
        self.project_manager.projects.append('Project E')
        with patch('json.dump'):
            result = self.project_manager.update_project_direct(
                'Project E', 'Project F')
            self.assertTrue(result)
            self.assertIn('Project F', self.project_manager.projects)
            self.assertNotIn('Project E', self.project_manager.projects)

    def tearDown(self):
        # Remove the test project file if it exists
        if os.path.exists(TEST_PROJECT_FILE):
            os.remove(TEST_PROJECT_FILE)


if __name__ == '__main__':
    unittest.main()
