# SteelSync Toolkit

[![Download SteelSync.exe](https://img.shields.io/badge/Download-SteelSync.exe%20(v1.1.0)-blue?style=for-the-badge&logo=windows)](https://github.com/DMan1906/SteelSync/releases/download/v1.1.0/SteelSync.exe)
[![Release](https://img.shields.io/github/v/release/DMan1906/SteelSync?style=for-the-badge)](https://github.com/DMan1906/SteelSync/releases/tag/v1.1.0)

> **Descarcă direct:** [👉 Click aici pentru SteelSync.exe (v1.1.0)](https://github.com/DMan1906/SteelSync/releases/download/v1.1.0/SteelSync.exe)
> 
> *Nu necesită instalare – se descarcă și se rulează direct.*

**SteelSync Toolkit** este o aplicație desktop pentru sortarea și extragerea fișierelor NC/NC1 destinate mașinilor de tăiat profile de oțel (FICEP, Voortman etc.).

## 📦 Descărcare și Rulare
1. Mergi la secțiunea **[Releases](https://github.com/DMan1906/SteelSync/releases)** din bara laterală dreapta (sau click pe butonul de mai sus).
2. Descarcă fișierul **`SteelSync.exe`** de la secțiunea **Assets**.
3. Rulează fișierul direct pe Windows.

## 🚀 Funcționalități

- **Sistem Auto-Update inteligent** – Verifică automat versiunile noi din GitHub Release la pornire (fără blocaje, 100% silențios când ești offline). Actualizare cu un singur click fără pierderea portabilității.
- **Reordonare Tab-uri & Personalizare UI** – Permite schimbarea ordinii tab-urilor principale prin Drag & Drop sau butoanele discrete `◀` / `▶` și meniu contextual (click-dreapta). Ordinea preferată se salvează automat și este reținută la repornire.
- **Memorii Independente pentru Browse** – Câmpul `Sursă` din `File Extractor` are memorie proprie izolată (`extractor_source_last_dir`), în timp ce toate celelalte butoane `Browse` partajează o memorie globală independentă (`global_last_dir`).
- **Design Complet Responsiv** – Adaptare dinamică la redimensionarea ferestrei și la ecrane compacte (`minsize: 900x650`). Câmpurile de introducere a căilor se extind proporțional pe orizontală, iar terminalul de loguri se scalează automat pe verticală.
- **File Extractor** – Caută și copiază fișierele unui proiect pe baza unei liste Excel (BOM).
- **NC Sorter** – Sortează fișierele .nc/.nc1 în foldere după tipul profilului:
  - PLĂCI (S235, S355, Z35, INOX, ALUMINIU)
  - PROFILE (HEA, HEB, IPE, UPN etc.)
  - BARE
  - CHESON (HSQEN, HSQ)
  - ȚEVI / T-URI / ȘINE
- **Modify NC** – Înlocuire text în masă în fișierele NC, ajustare automată a cantităților.

## 🏷️ Versiune curentă

**v1.1.0**

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
