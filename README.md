# SteelSync Toolkit

[![Download SteelSync.exe](https://img.shields.io/badge/Download-SteelSync.exe%20(v1.4.0)-blue?style=for-the-badge&logo=windows)](https://github.com/DMan1906/SteelSync/releases/download/v1.4.0/SteelSync.exe)
[![Release](https://img.shields.io/github/v/release/DMan1906/SteelSync?style=for-the-badge)](https://github.com/DMan1906/SteelSync/releases/tag/v1.4.0)

> **Descarcă direct:** [👉 Click aici pentru SteelSync.exe (v1.4.0)](https://github.com/DMan1906/SteelSync/releases/download/v1.4.0/SteelSync.exe)
> 
> *Nu necesită instalare – se descarcă și se rulează direct.*

**SteelSync Toolkit** este o aplicație desktop completă pentru industria construcțiilor metalice: căutare și extragere DWG-uri, sortarea și prelucrarea fișierelor NC/NC1 (FICEP, Voortman etc.) și managementul desenelor tehnice.

## 📦 Descărcare și Rulare
1. Mergi la secțiunea **[Releases](https://github.com/DMan1906/SteelSync/releases)** din bara laterală dreapta (sau click pe butonul de mai sus).
2. Descarcă fișierul **`SteelSync.exe`** de la secțiunea **Assets**.
3. Rulează fișierul direct pe Windows.

## 🚀 Funcționalități

- **🔍 DWG / Desene Finder (Extins în v1.4.0)** – Identifică și extrage automat desenele tehnice existente (`.dwg` și/sau `.pdf`) într-un arbore de clienți/proiecte:
  - **Suport Piese Sursă Flexibil (NOU în v1.4.0):** Posibilitate de alegere între fișiere piese **PDF (`.pdf`)**, **NC / NC1 (`.nc`, `.nc1`)** sau **Auto (Toate)**. Când ai doar fișiere NC/NC1, programul preia automat denumirile/codurile pieselor din acestea și caută desenele corespunzătoare în proiect.
  - **Format Desen Căutat în Proiect (NOU în v1.4.0):** Posibilitate de selecție între căutarea de **DWG (`.dwg`)**, **PDF (`.pdf`)** sau **Ambele (`.dwg` & `.pdf`)**.
  - **Păstrare subfoldere de grosimi:** Recunoaște și organizează automat fișierele în subfoldere de grosimi identice cu cele din proiect (`t10`, `t5`, `10`, `PL12` etc.).
  - **Suport complet pentru variații piese:** Trage automat toate variațiile aceluiași cod de piesă (ex: pentru `DSBO2014`, extrage `DSBO2014`, `DSBO2014T`, `DSBO2014C`, `DSBO2014L1`, excluzând fals-pozitivele numerice precum `DSBO20145`).
  - **Scanare recursivă completă:** Parcurge toate subdirectoarele până la capăt (`followlinks=True`), fără limite arbitrare de adâncime.
  - Tabel interactiv cu rezultate: **Desene Găsite** (deschidere fișier la dublu-click) și **Piese Fără Desen** (copiere rapidă a codurilor în clipboard).
  - Export raport detaliat Excel cu formatare profesională.
- **📁 Căutare Automată Proiect în Folderul Mamă** – Opțiune în `File Extractor` pentru scanarea automată a arborelui de clienți după număr/denumire proiect, folosind potrivire strictă pe cuvânt întreg (whole-word regex). Selectare automată când există o potrivire unică sau dialog modal intuitiv când există mai multe potriviri.
- **🖍️ Evidențiere Repere Lipsă cu Roșu în Rapoarte** – Colorează automat cu roșu (`FF9999`) rândurile pieselor care au fișiere lipsă în rapoartele Excel generate (`raport_PLACI.xlsx`, `raport_table.xlsx`, `raport_PROFILE.xlsx`), păstrând 100% intactă structura tabelelor fără a modifica numărul de rânduri sau coloane.
- **📊 Verificare Multi-Sheet la Rectificare Cantități NC** – Modulul `Modify NC` parcurge automat toate foile de calcul (sheets) din fișierul Excel de croi/BOM și cumulează corect cantitățile pieselor.
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

**v1.4.0**

## 📖 Ghid Dezvoltare și Reguli Versionare
Pentru ghidul complet de update și convențiile SemVer (`x.y.z`) dedicate dezvoltatorilor și agenților AI, consultă [AGENTS.md](AGENTS.md).

## Build

Aplicația este compilată cu PyInstaller:

```powershell
python -m PyInstaller SteelSync.spec --noconfirm
```
