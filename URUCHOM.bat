@echo off
chcp 65001 >nul
echo.
echo  Ewidencja Przebiegu Pojazdu - Smart Work v4.0
echo  ================================================
echo  BEZ LibreOffice - dziala od razu!
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [BLAD] Python nie jest zainstalowany!
    echo.
    echo  Pobierz ze strony: https://www.python.org/downloads/
    echo  Podczas instalacji zaznacz "Add Python to PATH"
    echo.
    pause & exit /b 1
)

python -c "import flask" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Instalowanie Flask i openpyxl...
    python -m pip install flask openpyxl --quiet
)

echo  Uruchamianie aplikacji...
echo  Otworzy sie przegladarka na: http://localhost:5000
echo.
start "" http://localhost:5000
python app.py
