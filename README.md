# SteelSync Toolkit

[![Download SteelSync.exe](https://img.shields.io/badge/Download-SteelSync.exe%20(v1.0.0)-blue?style=for-the-badge&logo=windows)](https://github.com/DMan1906/SteelSync/releases/download/v1.0.0/SteelSync.exe)
[![Release](https://img.shields.io/github/v/release/DMan1906/SteelSync?style=for-the-badge)](https://github.com/DMan1906/SteelSync/releases/tag/v1.0.0)

> **Descarcă direct:** [👉 Click aici pentru SteelSync.exe (v1.0.0)](https://github.com/DMan1906/SteelSync/releases/download/v1.0.0/SteelSync.exe)
> 
> *Nu necesită instalare – se descarcă și se rulează direct.*

**SteelSync Toolkit** este o aplicație desktop pentru sortarea și extragerea fișierelor NC/NC1 destinate mașinilor de tăiat profile de oțel (FICEP, Voortman etc.).

## 📦 Descărcare și Rulare
1. Mergi la secțiunea **[Releases](https://github.com/DMan1906/SteelSync/releases)** din bara laterală dreapta (sau click pe butonul de mai sus).
2. Descarcă fișierul **`SteelSync.exe`** de la secțiunea **Assets**.
3. Rulează fișierul direct pe Windows.


- **File Extractor** – Cauta si copiaza fisierele unui proiect pe baza unei liste Excel (BOM).
- **NC Sorter** – Sorteaza fisierele .nc/.nc1 in foldere dupa tipul profilului:
  - PLACI (S235, S355, Z35, INOX, ALUMINIU)
  - PROFILE (HEA, HEB, IPE, UPN, etc.)
  - BARE
  - CHESON (HSQEN, HSQ)
  - TEVI / T-URI / SINE
- **Modify NC** – Inlocuire text in masa in fisierele NC, ajustare cantitati.

## Versiune curenta

1.0.0

## Build

Aplicatia este compilata cu PyInstaller:

`
python -m PyInstaller SteelSync.spec --noconfirm
`

Executabilul rezultat se gaseste in dist/SteelSync.exe.

## Cerinte

- Python 3.10+
- customtkinter
- openpyxl
- pandas
