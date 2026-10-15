# SteelSync Toolkit

[![Download SteelSync.exe](https://img.shields.io/badge/Download-SteelSync.exe%20(v1.0.1)-blue?style=for-the-badge&logo=windows)](https://github.com/DMan1906/SteelSync/releases/download/v1.0.1/SteelSync.exe)
[![Release](https://img.shields.io/github/v/release/DMan1906/SteelSync?style=for-the-badge)](https://github.com/DMan1906/SteelSync/releases/tag/v1.0.1)

> **Descarcă direct:** [👉 Click aici pentru SteelSync.exe (v1.0.1)](https://github.com/DMan1906/SteelSync/releases/download/v1.0.1/SteelSync.exe)
> 
> *Nu necesită instalare – se descarcă și se rulează direct.*

**SteelSync Toolkit** este o aplicație desktop pentru sortarea și extragerea fișierelor NC/NC1 destinate mașinilor de tăiat profile de oțel (FICEP, Voortman etc.).

## 📦 Descărcare și Rulare
1. Mergi la secțiunea **[Releases](https://github.com/DMan1906/SteelSync/releases)** din bara laterală dreapta (sau click pe butonul de mai sus).
2. Descarcă fișierul **`SteelSync.exe`** de la secțiunea **Assets**.
3. Rulează fișierul direct pe Windows.

## 🚀 Funcționalități

- **Sistem Auto-Update inteligent** – Verifică automat versiunile noi din GitHub Release la pornire (fără blocaje, 100% silențios când ești offline). Actualizare cu un singur click fără pierderea portabilității.
- **File Extractor** – Caută și copiază fișierele unui proiect pe baza unei liste Excel (BOM).
- **NC Sorter** – Sortează fișierele .nc/.nc1 în foldere după tipul profilului:
  - PLĂCI (S235, S355, Z35, INOX, ALUMINIU)
  - PROFILE (HEA, HEB, IPE, UPN etc.)
  - BARE
  - CHESON (HSQEN, HSQ)
  - ȚEVI / T-URI / ȘINE
- **Modify NC** – Înlocuire text în masă în fișierele NC, ajustare automată a cantităților.

## 🏷️ Versiune curentă

**v1.0.1**

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
