#!/bin/bash
echo ""
echo " Ewidencja Przebiegu Pojazdu – Smart Work v4.0"
echo " ================================================"
echo " BEZ LibreOffice – działa od razu!"
echo ""
python3 -c "import flask" 2>/dev/null || pip3 install flask openpyxl --quiet
echo " Uruchamianie..."
sleep 1 && (open http://localhost:5000 2>/dev/null || xdg-open http://localhost:5000 2>/dev/null) &
python3 app.py
