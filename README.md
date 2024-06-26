# Outlook Meeting to Markdown

## Overview
This project's goal is to make it easier to pull Outlook meetings into Obsidian.

It includes a Python script that connects to Outlook and retrieves calendar events from the local machine. It provides summary information in markdown format for today's and tomorrow's meetings. The markdown includes a recommended wiki-style link for the meeting title as well as a bulleted list of all attendees, also in wiki-link format.

## Features
Retrieves today's and tomorrow's Outlook meetings.
Outputs meeting details in markdown format.
Generates wiki-style links for meeting titles and attendees.
Requirements
Windows operating system
Microsoft Outlook installed
Python 3.x
Required Python packages: art, pywin32

## Installation
1. Clone the repository:

```sh
git clone https://github.com/yourusername/outlook-meeting-to-markdown.git
cd outlook-meeting-to-markdown
```
2. Install the required Python packages:

```sh
pip install art pywin32
```
## Usage
1. Open a terminal or command prompt.
2. Navigate to the directory where the script is located.
3. Run the script:
```sh
python outlook_meeting_to_markdown.py
```
The script will print out the meetings for today and tomorrow, along with the details in markdown format.

## Using a Batch File
You can also use a batch file to run the script more conveniently. Below is an example of a batch file that activates a Python virtual environment and runs the script:

```batch
@echo off
echo Starting Outlook Meeting to Markdown

echo Script Directory: %SCRIPT_DIR%
set SCRIPT_DIR=%~dp0

echo Activating Python Virtual Environment
call %SCRIPT_DIR%venv\Scripts\activate.bat

echo Running Python Script
python %SCRIPT_DIR%outlook-meeting-to-markdown.py
pause
```

1. Save the above batch file with a .bat extension, for example, run_outlook_meetings.bat.
2. Place the batch file in the same directory as your Python script.
3. Double-click the batch file to run the script.

## Example Output
```lua
Subject: [[2023-06-16 Team Meeting]]
Start: 2023-06-16 09:00:00
Location: Conference Room A
Attendees:
  - "[[John Doe]]"
  - "[[Jane Smith]]"
  - "[[Bob Johnson]]"
----------------------------------------
Meeting Summary:
  - Subject: [[2023-06-16 Team Meeting]]
----------------------------------------
```

# Acknowledgments
- art
- pywin32