╔══════════════════════════════════════════════════════╗
║   EWIDENCJA PRZEBIEGU POJAZDU – Generator v4.0      ║
║   Smart Work Sp. z o.o.                              ║
╚══════════════════════════════════════════════════════╝

WYMAGANIA  (tylko 2 rzeczy!)
─────────────────────────────
• Python 3.8+  →  https://www.python.org/downloads/
  (zaznacz "Add Python to PATH" podczas instalacji!)

• To wszystko! LibreOffice NIE jest potrzebny.

URUCHOMIENIE
────────────
Windows:       Kliknij dwukrotnie  URUCHOM.bat
Mac / Linux:   chmod +x uruchom_mac_linux.sh
               ./uruchom_mac_linux.sh

Następnie otwórz przeglądarkę: http://localhost:5000
(otwiera się automatycznie)

JAK UŻYWAĆ
──────────
1. Wgraj plik GPS (.XLS lub .XLSX) – drag & drop lub kliknij
2. Wpisz imię i nazwisko kierowcy
3. Podaj stan licznika na początku miesiąca (km)
4. Opcjonalnie: dodaj daty tankowania (np. 5.12, 12.12)
5. Kliknij "Generuj" → plik Excel pobierze się automatycznie

WYGENEROWANY PLIK EXCEL
────────────────────────
Plik:  ewidencja_DSW3318E_2025_12.xlsx  (nazwa auta + miesiąc)

Arkusz 1 – Tytułowa:
  • Dane firmy Smart Work Sp. z o.o.
  • Nr rejestracyjny, marka i model pojazdu
  • Imię i nazwisko kierowcy
  • Stan licznika początku i końca miesiąca
  • Łączna liczba przejechanych km

Arkusz 2 – Rozlicznie:
  • Wszystkie 31 dni miesiąca
  • Opis trasy (adresy GPS: skąd → dokąd)
  • Kilometry każdego dnia
  • Imię i nazwisko kierowcy przy każdym wpisie
  • "tankowanie" w kolumnie Uwagi dla podanych dni

ZASADY AUTOMATYCZNE
───────────────────
  ✓ Sobota i Niedziela → km sumowane na poprzedni Piątek
  ✓ Dni bez jazdy pozostają puste
  ✓ Stan licznika końcowy = początkowy + suma km
