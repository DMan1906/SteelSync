# SteelSync Toolkit

[![Download SteelSync.exe](https://img.shields.io/badge/Download-SteelSync.exe%20(v1.7.7)-blue?style=for-the-badge&logo=windows)](https://github.com/DMan1906/SteelSync/releases/download/v1.7.7/SteelSync.exe)
[![Release](https://img.shields.io/github/v/release/DMan1906/SteelSync?style=for-the-badge)](https://github.com/DMan1906/SteelSync/releases/tag/v1.7.7)

> **Descarcă direct:** [👉 Click aici pentru SteelSync.exe (v1.7.7)](https://github.com/DMan1906/SteelSync/releases/download/v1.7.7/SteelSync.exe)
> 
> *Nu necesită instalare – se descarcă și se rulează direct.*

**SteelSync Toolkit** este o aplicație desktop completă pentru industria construcțiilor metalice: generare automată de tabele de import Metalix din rapoarte Word, organizare și regrupare avansată DWG, generare CAD 1:1 din desene PDF (fără cote), căutare și extragere DWG-uri, sortarea și prelucrarea fișierelor NC/NC1 (FICEP, Voortman etc.) și managementul desenelor tehnice.

## 📦 Descărcare și Rulare
1. Mergi la secțiunea **[Releases](https://github.com/DMan1906/SteelSync/releases)** din bara laterală dreapta (sau click pe butonul de mai sus).
2. Descarcă fișierul **`SteelSync.exe`** de la secțiunea **Assets**.
3. Rulează fișierul direct pe Windows.

## 🚀 Funcționalități

- **📐 PDF to DWG / CAD Generator 1:1 Fără Cote (PERFECȚIONAT în v1.7.7)**:
  - **Diferențiere Topologică Robustă între Decupaje/Notches și Rupturi Grafice:** Previne ștergerea sau unirea eronată a decupajelor semicirculare sau curbe (ex. `R35`), verificând continuitatea topologică a nodurilor și conservând toate detaliile geometrice ale conturului.
  - **Calibrare Matematică Exactă a Pauzelor Multi-Break cu Decupaje:** Garantează cote 100% exacte la fiecare segment al piesei (ex. 2353, 100, 1100, 100, 1100, 100, 1200, 100, 612 mm). Decupajele dintre stațiile de ruptură sunt izolate complet și nu mai afectează calculul expansiunilor.
  - **Detecție și Export Nativ de Găuri (`CIRCLE`):** Detectează automat toate găurile circulare ale piesei (inclusiv găuri filetate sau standard) și le generează ca entități native `CIRCLE` pe layer-ul `0`, poziționate exact conform cotelor de execuție.
  - **Corecție Multi-Break pentru Piese Plate/Profiluri:** Calculează expansiunile corecte peste multiple stații de ruptură Tekla fără a depinde de prezența liniilor verticale intermediare, distribuind lungimile exact după cotele intermediare ale găurilor.
  - **Extragere Robustă LENGDE prin Aliniere Spațială:** Extrage lungimea exactă a piesei din coloana `LENGDE` a tabelelor STYKKLISTE, eliminând complet confuzia cu numărul de bucăți (`ANTALL`), numărul de rând sau greutatea.
  - **Bridging Garantat pentru Rupturi Tekla:** Orice pauză/ruptură grafică Tekla (shortened view) este unită fără goluri, extinzând piesa la lungimea reală 1:1 sau unind marginile pentru un contur închis garantat.
  - **Generare Nativă în Layer-ul Default (`0`):** Geometria piesei este creată direct în layer-ul nativ `0` (fără layere suplimentare create de program), conform cerințelor standard din producție și AutoCAD.
  - **Reconstrucție Topologică cu Snapping Magnetic (Grid & Union-Find):** Elimină 100% micro-golurile dintre linii (gaps de 0.01 - 0.5 mm) și treptele decalate, unind nodurile adiacente în puncte identice exacte.
  - **Generare Nativă de Polilinii Închise (`LWPOLYLINE`):** Conturul piesei și toate decupajele/găurile interioare sunt generate ca entități polilinie continuă închisă (`closed=True`), gata pentru import direct în softurile de nesting (Metalix, Lantek, Pronest) fără erori de "contur deschis".
  - **Configurare Nativă Stil Cotare 1:1 (`DIMLFAC = 1.0`):** Toate fișierele DWG/DXF generate au stilul de cotare presetat cu factorul de multiplicare 1.0 (nu 100.0), astfel încât cotele puse în AutoCAD (`DIMENSION`) să corespundă 100% cu măsurătoarea reală (`MEASURE`).
  - **Racordări Curbe Native (Fillet Arcs via Bulge):** Arcele de racordare sunt încorporate matematic direct în polilinie prin parametrul `bulge`, eliminând aproximarea prin linii frânte sau arce decuplate.
  - **Filtrare Avansată Cote și Săgeți de Secțiune:** Ignoră săgețile mari de secțiune (A-A, B-B), liniile ajutătoare de cotă și cartușul, izolând strict conturul piesei chiar și atunci când tabelul STYKKLISTE se află în partea superioară a desenului.
  - **Suport Multi-Break Tekla pentru Decupaje Marginale:** Piesele lungi scurtate în Tekla cu decupaje deschise (notches) își păstrează forma corectă cu golurile decupate din contur (fără linii suprapuse prin decupaje).

- **📑 Imports Generator (NOU în v1.7.0 - Pagină / Tab Dedicat)**:
  - Modul specializat pentru procesarea rapoartelor de nesting Metalix salvate în format Word (`.doc` și `.docx`).
  - Extrage automat tabelele de piese din documente: denumire piesă, grosime tablă, calitate material (`S235`, `S355`, etc.), cantitate totală necesară și denumire nesting.
  - Generare fișier Excel standardizat (`IMPORT.xlsx`) gata de importat direct în softul de nesting Metalix.
  - Funcție **Quick Save on Desktop** printr-un singur click (`IMPORT.xlsx` generat direct pe desktop) și buton dedicat **Copiază în Clipboard** (date tabelare compatibile direct cu Paste în orice spreadsheet Excel).
  - Previzualizare tabelară interactivă cu căutare/filtrare live a pieselor extrase.

- **📦 Integrare DWG Tools în DWG Finder (NOU în v1.7.0)**:
  - Tabul `🔍 DWG Finder` integrează acum un comutator vizual rapid de sub-vederi (segmented button) cu 3 moduri de lucru:
    1. **🔍 Căutare DWG în Proiecte:** Modulul consacrat de scanare recursivă a desenelor DWG/PDF după codurile pieselor sau fișiere NC.
    2. **📦 Organizare Piese (BOM ➜ Piese):** Citește lista de debitare/BOM din Excel, detectează automat coloanele (Piesă, Grosime/Profil, Material, Cantitate), caută fișierele DWG în folderul sursă și le organizează automat în foldere curate per grosime (ex. `t10`, `t15`, `t20`). Include rapoarte automate `_PIESE_LIPSA.txt` și `_RAPORT_ORGANIZARE.txt`.
    3. **🔄 Regrupare DWG pe Grosimi:** Scanează o ierarhie dezorganizată de fișiere DWG, identifică automat grosimea fiecărui reper din denumirea fișierului sau din calea folderului părinte și le re-grupează curat în subfoldere per grosime.

- **📐 PDF to DWG / CAD Generator 1:1 Fără Cote (v1.6.0)**:
  - Modul dedicat și tab nou în aplicație: convertește automat desenele tehnice de execuție PDF în fișiere CAD (`.dwg` / `.dxf`) la scară reală **1:1 (milimetri)**.
  - **Filtrare Geometrie Pură:** Elimină automat toate cotele, liniile de cotă, săgețile, textele, simbolurile de secțiune și cartușul desenului, păstrând doar conturul piesei și decupajele interioare.
  - **Calibrare Automată:** Extrage automat metadatele din tabelul `STYKKLISTE` (dimensiuni nominale profil/lungime, scară) și calibrează geometria la milimetru, mutând colțul stânga-jos în originea CAD `(0, 0)`.
  - Suport **Single (Fișier Unic)** și **Batch (Folder Întreg de PDF-uri)**.
  - **Generare DWG directă din `DWG Finder`:** Piesele identificate ca lipsă pot fi transformate direct în DWG-uri curate printr-un singur click pe butonul `⚡ Generează DWG pentru Piesele Lipsă (din PDF)`.
  - Suport nativ pentru `.dxf` (R2010, compatibil cu orice versiune de AutoCAD/CAD) și export `.dwg` nativ prin integrare directă cu ODA File Converter.
- **⚡ Căutare Rapidă și Sigură a Folderului de Proiect (v1.6.0)**:
  - Căutarea după `Nr. / Denumire Proiect` rulează complet asincron pe un thread dedicat de fundal, eliminând orice înghețare a interfeței pe drive-uri de rețea (`Q:\...`).
  - Scanare inteligentă pe structura specifică `Folder Mamă -> Client -> Proiect`: localizează directorul proiectului fără a intra niciodată în interiorul fișierelor sau subfolderelor de proiect, oferind rezultate aproape instantanee.
- **📑 Extracție Desene de Ansamblu din Excel Ansambluri (v1.5.0)**:
  - Opțiune dedicată în `File Extractor` pentru parsarea automată a fișierelor Excel cu ansambluri (Assembly Part List, ex. `1242 - 48.2 Fiskebrygga A4.xlsx`).
  - Auto-detecție inteligentă a coloanei de ansamblu (`Denumire ansamblu` / Coloana B), ignorând sub-piesele și rândurile de sumar.
  - Căutare recursivă cu toleranță la cratimă/revizii și salvare automată în subfolderul dedicat `Destinație/Desene Ansamblu`.
  - Generare raport detaliat `Raport_Status_Ansambluri.xlsx` cu evidențierea stadiului fiecărui ansamblu (GĂSIT / LIPSĂ).
- **⚡ Optimizare Redimensionare și UI Fluid**:
  - Debounce inteligent la resize pentru o experiență 60+ FPS complet lipsită de sacadare/lag la tragerea marginilor ferestrei.
  - Layout pe rânduri fluid pentru `Extensii Fișiere`: toate extensiile și butonul `+ Add` încap impecabil pe o singură linie și se pliază/depliază natural fără goluri sau butoane tăiate.
  - Design îmbunătățit al barei de navigare/tab-uri fără artefacte vizuale sau halo-uri gri decalate.
- **🔍 DWG / Desene Finder** – Identifică și extrage automat desenele tehnice existente (`.dwg` și/sau `.pdf`) într-un arbore de clienți/proiecte:
  - **Suport Piese Sursă Flexibil:** Posibilitate de alegere între fișiere piese **PDF (`.pdf`)**, **NC / NC1 (`.nc`, `.nc1`)** sau **Auto (Toate)**. Când ai doar fișiere NC/NC1, programul preia automat denumirile/codurile pieselor din acestea și caută desenele corespunzătoare în proiect.
  - **Format Desen Căutat în Proiect:** Posibilitate de selecție între căutarea de **DWG (`.dwg`)**, **PDF (`.pdf`)** sau **Ambele (`.dwg` & `.pdf`)**.
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
- **File Extractor** – Caută și copiază fișierele unui proiect pe baza unei liste Excel (BOM).
- **NC Sorter** – Sortează fișierele .nc/.nc1 în foldere după tipul profilului:
  - PLĂCI (S235, S355, Z35, INOX, ALUMINIU)
  - PROFILE (HEA, HEB, IPE, UPN etc.)
  - BARE
