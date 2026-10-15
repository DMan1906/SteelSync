# SteelSync Toolkit

**SteelSync Toolkit** este o aplicatie desktop pentru sortarea si extragerea fisierelor NC/NC1 destinate masinilor de taiat profile de otel (FICEP, Voortman, etc.).

## Functionalitati

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
