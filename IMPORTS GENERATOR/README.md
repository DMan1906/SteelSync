# Generator Import-uri Excel din Word (Metalix `.DOC` / `.DOCX` ➜ `.XLSX`)

Aplicație desktop pentru extragerea automată a tabelului **Parts** din rapoartele Word generate de Metalix (`.DOC` sau `.DOCX`) și generarea fișierelor Excel de import (`.XLSX`).

## Reguli de curățare a denumirii piesei (Coloana A în Excel)
1. **Eliminare extensie**: `.dft`, `.DFT`, `.dxf`, `.dwg`, `.nc`, `.nc1` sunt șterse automat.
2. **Eliminare sufixe de indexare Metalix**: `_001`, `_002`, `_003` etc. sunt șterse automat (inclusiv când apar înainte de `+1.1`, ex: `DSB2419L_001+1.1.DFT` ➜ `DSB2419L+1.1`).
3. **Eliminare caractere speciale**: cratimele `-`, underscore `_` și spațiile sunt eliminate (ex: `DSB-2259.dft` ➜ `DSB2259`).
4. **Păstrare excepții `+1.1`, `+1.2`, `+1.3` etc.**: sufixele de forma `+1.1`, `+1.2`, `+1.3` etc. sunt păstrate intacte (ex: `DSB2052L+1.1.DFT` ➜ `DSB2052L+1.1`).

## Exemple de transformare
| Denumire în Word (`Part name`) | Denumire în Excel (`Col A`) | Cantitate (`Col B`) |
| :--- | :--- | :---: |
| `DSB-2259.dft` | `DSB2259` | `2` |
| `DSP-2037.dft` | `DSP2037` | `1` |
| `DSB2273L1_001.dft` | `DSB2273L1` | `1` |
| `DSB2046T_002.dft` | `DSB2046T` | `1` |
| `DSB2052L+1.1.DFT` | `DSB2052L+1.1` | `2` |
| `DSB2419L_001+1.1.DFT` | `DSB2419L+1.1` | `1` |

## Pornire
- Dublu-click pe **`GeneratorImportExcel.exe`** sau pe **`Start_Aplicatie.bat`** (sau `python app.py`).

