from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QListWidget, QLineEdit, QPushButton
import json
import os

PROJECTS_FILE = 'projects.json'

class ProjectManager:
    def __init__(self, parent):
        self.parent = parent
        self.meeting_manager = None
        self.projects = self.load_projects()
        if "Unassigned" not in self.projects:
            self.projects.insert(0, "Unassigned")  # Ensure "Unassigned" is always the first project
        self.layout = QHBoxLayout()

        self.project_list_widget = QListWidget()
        self.project_list_widget.setSelectionMode(QListWidget.SingleSelection)
        self.layout.addWidget(self.project_list_widget)

        self.buttons_layout = QVBoxLayout()

        self.project_entry = QLineEdit()
        self.buttons_layout.addWidget(self.project_entry)

        add_project_button = QPushButton("Add Project")
        add_project_button.clicked.connect(self.add_project)
        self.buttons_layout.addWidget(add_project_button)

        remove_project_button = QPushButton("Remove Project")
        remove_project_button.clicked.connect(self.remove_project)
        self.buttons_layout.addWidget(remove_project_button)

        update_project_button = QPushButton("Update Project")
        update_project_button.clicked.connect(self.update_project)
        self.buttons_layout.addWidget(update_project_button)

        self.layout.addLayout(self.buttons_layout)

        self.refresh_project_list()

    def set_meeting_manager(self, meeting_manager):
        self.meeting_manager = meeting_manager
        self.refresh_project_list()

    def load_projects(self):
        if os.path.exists(PROJECTS_FILE):
            with open(PROJECTS_FILE, 'r') as file:
                return json.load(file)
        return ['Project A', 'Project B', 'Project C']

    def save_projects(self):
        with open(PROJECTS_FILE, 'w') as file:
            json.dump(self.projects, file)

    def refresh_project_list(self):
        self.project_list_widget.clear()
        for project in self.projects:
            if project != "Unassigned":  # Do not show "Unassigned" in the list
                self.project_list_widget.addItem(project)
        if self.meeting_manager:
            self.meeting_manager.refresh_dropdowns()

    def add_project(self):
        new_project = self.project_entry.text()
        if new_project and new_project not in self.projects:
            self.projects.append(new_project)
            self.save_projects()
            self.project_entry.clear()
            self.refresh_project_list()
            self.parent.notification_label.setText(f"Project '{new_project}' has been added.")
        else:
            self.parent.notification_label.setText("Please enter a valid and unique project name.")

    def remove_project(self):
        selected_project = self.project_list_widget.currentItem()
        if selected_project and selected_project.text() in self.projects and selected_project.text() != "Unassigned":
            self.projects.remove(selected_project.text())
            self.save_projects()
            self.refresh_project_list()
            self.parent.notification_label.setText(f"Project '{selected_project.text()}' has been removed.")
        else:
            self.parent.notification_label.setText("Please select a valid project to remove.")

    def update_project(self):
        selected_project_index = self.project_list_widget.currentRow()
        if selected_project_index >= 0:
            selected_project = self.project_list_widget.item(selected_project_index).text()
            new_name = self.project_entry.text()
            if new_name and new_name not in self.projects:
                self.projects[selected_project_index + 1] = new_name  # Adjust index to skip "Unassigned"
                self.save_projects()
                self.project_entry.clear()
                self.refresh_project_list()
                self.parent.notification_label.setText(f"Project '{selected_project}' has been renamed to '{new_name}'.")
            else:
                self.parent.notification_label.setText("Please enter a valid and unique project name.")
        else:
            self.parent.notification_label.setText("Please select a project to update.")
