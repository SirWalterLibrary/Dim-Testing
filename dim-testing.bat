@echo off
setlocal

:: Navigate to the project folder
cd %~dp0\Dim-Testing

:: Activate the virtual environment
call venv\Scripts\activate

:: Run the Python script
python dim-testing.py

:: Keep the command prompt open if needed (optional)
cmd /k
