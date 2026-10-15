# 🤖 Ghid de Dezvoltare și Proces de Update pentru Agenți AI (SteelSync Toolkit)

Acest document stabilește regulile obligatorii, convențiile de versionare și procedura pas cu pas pe care orice agent AI (sau dezvoltator) **trebuie** să le urmeze atunci când aduce modificări, corecții sau noi funcționalități aplicației **SteelSync Toolkit**.

---

## 📌 1. Regula de Versionare SemVer (x.y.z)

Versiunea aplicației este definită strict în formatul `x.y.z`:
* **`z` (Patch / Minor bug fix):** Se incrementează pentru update-uri mici, rezolvări de bug-uri, corecții de text/UI, ajustări de permisiuni sau optimizări de cod care nu adaugă funcționalități noi.
  * *Exemplu:* `1.0.1` -> `1.0.2`
* **`y` (Minor / Feature):** Se incrementează pentru update-uri medii și funcționale (ex: adăugarea unui tab nou, adăugarea de noi tipuri de sortare NC, integrare cu noi formate Excel/BOM, funcții noi în File Extractor sau Modify NC).
  * *Exemplu:* `1.0.2` -> `1.1.0`
* **`x` (Major / Breaking change):** Se incrementează pentru update-uri majore de arhitectură, rescrierea interfeței grafice sau schimbări fundamentale de compatibilitate.
  * *Exemplu:* `1.1.0` -> `2.0.0`

> ⚠️ **Notă:** Nu folosiți sufixe temporare (cum ar fi `1.0.1t`) în producție; versiunile trebuie să fie strict SemVer standard (`x.y.z`).

---

## 🔄 2. Cum Funcționează Sistemul de Auto-Update

1. **Verificare silențioasă și asincronă (Offline-first):**
   * La 1.5 secunde după inițializarea ferestrei, un thread separat (`daemon=True`) interoghează GitHub API:
     `https://api.github.com/repos/DMan1906/SteelSync/releases/latest`
   * Timeout-ul este scurt (3.5 – 4 secunde). Dacă utilizatorul nu are internet sau GitHub nu răspunde, aplicația rulează 100% normal, fără nicio notificare de eroare.
2. **Comparare versiuni:**
   * Compară `tag_name` din GitHub Release (ex. `v1.0.2`) cu `APP_VERSION` din cod.
   * Dacă `tag_name` este mai nou și nu se află în `ignored_update_version` din `steelsync_settings.json`, se afișează dialogul `UpdateDialog`.
3. **Descărcare & Înlocuire pe loc (Hot-Swap în Windows):**
   * Se descarcă noul `SteelSync.exe` în `%TEMP%\SteelSync_update_<tag>.exe`.
   * Se generează scriptul `%TEMP%\steelsync_updater.bat`.
   * **Regulă critică PyInstaller (`_MEIPASS2`):** Variabila de mediu `_MEIPASS2` și `_MEIPASS` **trebuie obligatoriu curățate** (`set _MEIPASS2=` și eliminarea din `env` în Python). Altfel, noua aplicație moștenește folderul temporar al vechiului proces deja șters și crapă cu eroarea `Failed to load Python DLL`.
   * **Regulă critică de sincronizare:** Scriptul `.bat` așteaptă oprirea completă a tuturor proceselor `SteelSync.exe` și include o pauză de siguranță de **2 secunde** pentru ca PyInstaller să își șteargă folderul temporar `_MEI...` înainte de copiere.
   * Scriptul copiază noul executabil peste cel vechi, setează directorul de lucru (`cd /d "%TARGET_DIR%"`) și repornește noua versiune curat.

---

## 🛠️ 3. Procedura Pas-cu-Pas după Fiecare Modificare de Cod

De fiecare dată când agentul aduce un update proiectului, **TREBUIE** parcurși următorii pași în ordine:

### Pasul 1: Incrementarea Versiunii în Cod
În fișierul `part_extractor.py`, actualizează constanta globală:
```python
APP_VERSION = "x.y.z"  # Incrementează conform regulii SemVer
```

### Pasul 2: Actualizarea README.md
Actualizează insigna și link-ul de descărcare din `README.md` la noua versiune `vX.Y.Z`.

### Pasul 3: Validarea Sintaxei
Verifică integritatea codului înainte de compilare:
```powershell
python -m py_compile part_extractor.py
```

### Pasul 4: Recompilarea Executabilului (Build curat)
Compilează noul binar `SteelSync.exe` folosind fișierul `.spec`:
```powershell
python -m PyInstaller SteelSync.spec --noconfirm
```
*(Executabilul rezultat se va afla în `dist\SteelSync.exe`).*

### Pasul 5: Testarea Rulării Executabilului
Verifică printr-un test scurt că `dist\SteelSync.exe` pornește fără erori:
```powershell
$p = Start-Process -FilePath "dist\SteelSync.exe" -PassThru
Start-Sleep -Seconds 3
if (-not $p.HasExited) { $p.Kill() }
```

### Pasul 6: Commit și Push în Git
Sincronizează codul sursă în repository:
```powershell
git add part_extractor.py README.md AGENTS.md
git commit -m "feat/fix: descriere modificari (vX.Y.Z)"
git push origin main
```

### Pasul 7: Crearea Release-ului pe GitHub (Obligatoriu pentru Auto-Update)
Auto-update-ul funcționează **doar dacă** noul `.exe` este publicat ca GitHub Release:
```powershell
$env:PATH = $env:PATH + ";C:\Program Files\GitHub CLI"
gh release create vX.Y.Z "dist\SteelSync.exe#SteelSync.exe" --repo DMan1906/SteelSync --title "vX.Y.Z - Titlu Release" --notes "Note de lansare si descrierea functiilor noi..."
```

### Pasul 8: Actualizarea Executabilului de pe Desktop
Dacă pe desktop-ul utilizatorului există `SteelSync.exe`, copiază noua versiune pentru a asigura sincronizarea locală:
```powershell
Copy-Item -Path "dist\SteelSync.exe" -Destination "C:\Users\Harmat Vajda\Desktop\SteelSync.exe" -Force
```

---

## 📋 4. Fișiere Cheie ale Proiectului

* `part_extractor.py` – Codul sursă principal (GUI CustomTkinter, File Extractor, NC Sorter, Modify NC, Auto-Updater).
* `nc_sorter.py` – Modulul de logică pentru sortarea și parsarea fișierelor NC/NC1.
* `SteelSync.spec` – Configurația PyInstaller pentru generarea unui singur executabil portabil (`--onefile`, iconiță, librării CustomTkinter).
* `app_icon.ico` – Iconița oficială a aplicației.
* `README.md` – Documentația publică a repository-ului și butonul de descărcare directă.
* `AGENTS.md` – Prezentul ghid pentru agenți AI.
