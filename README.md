# SteelSync Toolkit

[![Download SteelSync.exe](https://img.shields.io/badge/Download-SteelSync.exe%20(v1.2.0)-blue?style=for-the-badge&logo=windows)](https://github.com/DMan1906/SteelSync/releases/download/v1.2.0/SteelSync.exe)
[![Release](https://img.shields.io/github/v/release/DMan1906/SteelSync?style=for-the-badge)](https://github.com/DMan1906/SteelSync/releases/tag/v1.2.0)

> **Descarcă direct:** [👉 Click aici pentru SteelSync.exe (v1.2.0)](https://github.com/DMan1906/SteelSync/releases/download/v1.2.0/SteelSync.exe)
> 
> *Nu necesită instalare – se descarcă și se rulează direct.*

**SteelSync Toolkit** este o aplicație desktop completă pentru industria construcțiilor metalice: căutare și extragere DWG-uri, sortarea și prelucrarea fișierelor NC/NC1 (FICEP, Voortman etc.) și managementul desenelor tehnice.

## 📦 Descărcare și Rulare
1. Mergi la secțiunea **[Releases](https://github.com/DMan1906/SteelSync/releases)** din bara laterală dreapta (sau click pe butonul de mai sus).
2. Descarcă fișierul **`SteelSync.exe`** de la secțiunea **Assets**.
3. Rulează fișierul direct pe Windows.

## 🚀 Funcționalități

- **🔍 DWG Finder (NOU în v1.2.0)** – Identifică și extrage automat fișierele CAD `.dwg` existente într-un arbore de clienți/proiecte pe baza unui pachet de fișiere `.pdf` cu piese de desenat:
  - Localizează automat folderele proiectului după număr în orice structură de directoare.
  - Algoritm avansat de normalizare și potrivire cod piesă (elimină caractere speciale, prefixe de proiect, paranteze și revizii, fără fals-pozitive pe cifre).
  - Fallback inteligent bazat pe textul extras din antetul PDF (`pypdf`).
  - Tabel interactiv cu rezultate: **Piese cu DWG Găsit** (deschidere fișier la dublu-click) și **Piese Fără DWG (Lipsă)** (copiere rapidă a codurilor în clipboard).
  - Export raport detaliat Excel cu formatare profesională pe două foi de lucru.
- **Sistem Auto-Update inteligent** – Verifică automat versiunile noi din GitHub Release la pornire (fără blocaje, 100% silențios când ești offline). Actualizare cu un singur click fără pierderea portabilității.
- **Reordonare Tab-uri Floating prin Drag & Drop** – Tab-urile își păstrează estetica nativă flotantă (floating pills) și pot fi reordonate direct prin Drag & Drop cu mouse-ul (sau prin click-dreapta / scurtături). Ordinea preferată se salvează automat și este reținută la repornire.
- **Memorii Independente pentru Browse** – Câmpul `Sursă` din `File Extractor` are memorie proprie izolată (`extractor_source_last_dir`), în timp ce toate celelalte butoane `Browse` partajează o memorie globală independentă (`global_last_dir`).
- **Design Complet Responsiv & Randare Ultra-Fluidă** – Fără sacadări la redimensionarea ferestrei (motor optimizat de randare nativă și debouncing). Elementele din cardul de extensii se adaptează dinamic și urcă pe rândurile superioare când au spațiu disponibil, iar câmpurile de input se extind proporțional (`minsize: 900x650`).
- **File Extractor** – Caută și copiază fișierele unui proiect pe baza unei liste Excel (BOM).
- **NC Sorter** – Sortează fișierele .nc/.nc1 în foldere după tipul profilului:
  - PLĂCI (S235, S355, Z35, INOX, ALUMINIU)
  - PROFILE (HEA, HEB, IPE, UPN etc.)
  - BARE
  - CHESON (HSQEN, HSQ)
  - ȚEVI / T-URI / ȘINE
- **Modify NC** – Înlocuire text în masă în fișierele NC, ajustare automată a cantităților.

## 🏷️ Versiune curentă

**v1.2.0**

## 📖 Ghid Dezvoltare și Reguli Versionare
Pentru ghidul complet de update și convențiile SemVer (`x.y.z`) dedicate dezvoltatorilor și agenților AI, consultă [AGENTS.md](AGENTS.md).

## Build

Aplicația este compilată cu PyInstaller:

```powershell
python -m PyInstaller SteelSync.spec --noconfirm
```

Executabilul rezultat se găsește în `dist/SteelSync.exe`.

## Cerințe

- Python 3.10+
- customtkinter
- openpyxl
- pandas
- pypdf
