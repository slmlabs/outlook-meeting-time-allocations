# Outlook Meeting Time Allocations App

## Overview

The Meeting Assignment Application helps users manage their Outlook meetings by allowing them to assign meetings to specific projects and generate reports on time spent per project.

## Features

- Fetch meetings from Outlook within a specified date range.
- Assign meetings to projects.
- Generate reports on time spent per project.
- Check if Outlook is running and provide options to start Outlook if not running.

## Installation

1. **Clone the repository:**

   ```sh
   git clone https://github.com/slmlabs/outlook-meeting-time-allocations.git
   cd outlook-meeting-time-allocations
   ```

2. **Create a virtual environment and activate it:**

   ```sh
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install the dependencies:**
   ```sh
   pip install -r requirements.txt
   ```

## Usage

1. **Run the application:**

   ```sh
   python main.py
   ```

2. **Follow the on-screen instructions:**
   - Select the date range to fetch meetings.
   - Assign projects to meetings using the provided dropdown.
   - Generate reports to view time spent per project.

## Testing

Run the tests to ensure everything is working correctly:

```sh
python -m unittest discover -s tests
```

## Contributing

1. **Fork the repository.**

2. **Create a new branch for your feature or bug fix:**

   ```sh
   git checkout -b feature-name
   ```

3. **Make your changes and commit them:**

   ```sh
   git commit -m "Description of your changes"
   ```

4. **Push to your branch:**

   ```sh
   git push origin feature-name
   ```

5. **Create a pull request.**

## License

This project is licensed under the MIT License.
