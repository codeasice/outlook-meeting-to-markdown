@echo off
echo Starting Outlook Meeting to Markdown

echo Script Directory: %SCRIPT_DIR%
set SCRIPT_DIR=%~dp0

echo Activating Python Virtual Environment
call %SCRIPT_DIR%venv\Scripts\activate.bat

echo Running Python Script
python %SCRIPT_DIR%outlook-meeting-to-markdown.py
pause