import json
import os
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QListWidget, QLineEdit, QPushButton, QWidget, QLabel

PROJECTS_FILE = 'projects.json'


class ProjectManager(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.meeting_manager = None
        self.projects = self.load_projects()
        self.ensure_unassigned_project()
        self.setup_ui()

    def setup_ui(self):
        """Set up the user interface."""
        main_layout = QHBoxLayout(self)

        self.project_list_widget = QListWidget()
        self.project_list_widget.setSelectionMode(QListWidget.SingleSelection)
        main_layout.addWidget(self.project_list_widget)

        self.buttons_layout = QVBoxLayout()
        main_layout.addLayout(self.buttons_layout)

        project_entry_layout = QHBoxLayout()
        self.project_entry_label = QLabel("Project Name:")
        project_entry_layout.addWidget(self.project_entry_label)

        self.project_entry = QLineEdit()
        project_entry_layout.addWidget(self.project_entry)

        self.buttons_layout.addLayout(project_entry_layout)

        self.add_project_button = self.create_button(
            "Add Project", self.add_project)
        self.buttons_layout.addWidget(self.add_project_button)

        self.remove_project_button = self.create_button(
            "Remove Project", self.remove_project)
        self.buttons_layout.addWidget(self.remove_project_button)

        self.update_project_button = self.create_button(
            "Update Project", self.update_project)
        self.buttons_layout.addWidget(self.update_project_button)

        self.setLayout(main_layout)
        self.refresh_project_list()

    def create_button(self, text, callback):
        """Create a button with the given text and callback."""
        button = QPushButton(text)
        button.clicked.connect(callback)
        return button

    def ensure_unassigned_project(self):
        """Ensure 'Unassigned' is always in the project list."""
        if "Unassigned" not in self.projects:
            self.projects.insert(0, "Unassigned")

    def set_meeting_manager(self, meeting_manager):
        """Set the meeting manager and refresh the project list."""
        self.meeting_manager = meeting_manager
        self.refresh_project_list()

    def load_projects(self):
        """Load projects from a JSON file."""
        if os.path.exists(PROJECTS_FILE):
            with open(PROJECTS_FILE, 'r') as file:
                return json.load(file)
        return ['Project A', 'Project B']

    def save_projects(self):
        """Save projects to a JSON file."""
        with open(PROJECTS_FILE, 'w') as file:
            json.dump(self.projects, file)

    def refresh_project_list(self):
        """Refresh the project list in the UI."""
        self.project_list_widget.clear()
        for project in self.projects:
            if project != "Unassigned":  # Do not show "Unassigned" in the list
                self.project_list_widget.addItem(project)
        if self.meeting_manager:
            self.meeting_manager.refresh_dropdowns()

    def add_project(self):
        """Add a new project."""
        new_project = self.project_entry.text().strip()
        if new_project:
            if self.add_project_direct(new_project):
                self.project_entry.clear()
                self.refresh_project_list()
        else:
            self.show_notification("Please enter a valid project name.")

    def remove_project(self):
        """Remove the selected project."""
        selected_project = self.project_list_widget.currentItem()
        if selected_project:
            project_name = selected_project.text()
            if self.remove_project_direct(project_name):
                self.refresh_project_list()
        else:
            self.show_notification("Please select a project to remove.")

    def update_project(self):
        """Update the selected project with a new name."""
        selected_project = self.project_list_widget.currentItem()
        new_name = self.project_entry.text().strip()
        if selected_project and new_name:
            project_name = selected_project.text()
            if self.update_project_direct(project_name, new_name):
                self.project_entry.clear()
                self.refresh_project_list()
        else:
            self.show_notification(
                "Please select a project and enter a new name.")

    def add_project_direct(self, new_project):
        """Add a new project to the list and save."""
        if new_project not in self.projects:
            self.projects.append(new_project)
            self.save_projects()
            self.show_notification(f"Project '{new_project}' has been added.")
            return True
        self.show_notification("Project name must be unique.")
        return False

    def remove_project_direct(self, project_name):
        """Remove a project from the list and save."""
        if project_name in self.projects and project_name != "Unassigned":
            self.projects.remove(project_name)
            self.save_projects()
            self.show_notification(
                f"Project '{project_name}' has been removed.")
            return True
        self.show_notification(
            "Cannot remove 'Unassigned' or non-existent project.")
        return False

    def update_project_direct(self, old_name, new_name):
        """Update a project's name and save."""
        if old_name in self.projects and new_name not in self.projects and old_name != "Unassigned":
            index = self.projects.index(old_name)
            self.projects[index] = new_name
            self.save_projects()
            self.show_notification(
                f"Project '{old_name}' has been renamed to '{new_name}'.")
            return True
        self.show_notification("Please enter a valid and unique project name.")
        return False

    def show_notification(self, message):
        """Show a notification message."""
        self.parent.notification_label.setText(message)
