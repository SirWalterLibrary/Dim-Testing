@echo off
setlocal enabledelayedexpansion

:: Define repository URL and folder name
set REPO_URL=https://github.com/SirWalterLibrary/Dim-Testing.git
set FOLDER_NAME=Dim-Testing

:: Clone the repository
echo Cloning repository...
git clone %REPO_URL%

:: Check if cloning was successful
if not exist %FOLDER_NAME% (
    echo Error: Failed to clone repository.
    exit /b 1
)

:: Change directory to the cloned repository
cd %FOLDER_NAME%

:: Create virtual environment
echo Creating virtual environment...
python -m venv venv

:: Activate the virtual environment
echo Activating virtual environment...
call venv\Scripts\activate

:: Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

echo Setup complete. Virtual environment is activated.
echo You can now run your Python scripts within this environment.

:: Keep the command prompt open
cmd /k
