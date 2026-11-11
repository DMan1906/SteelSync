# 📁 Mini Extragere & Organizare NC, PDF & DWG

Aplicație desktop completă dedicată pentru industria construcțiilor metalice și prelucrarea reperelor CAD/CAM (Tekla Structures, Advance Steel, Ficep, Voortman etc.).

Aplicația oferă două module puternice accesibile prin tab-uri dedicate:

---

## 📦 Modulul 1: Extragere pe Piese (BOM Excel -> Piese)

1. **Citește tabelul Excel**:
   - Detectează automat foaia de calcul și rândul de antet.
   - Detectează automat coloana cu **Denumirea Ansamblului** și coloana cu **Denumirea Piesei**.
   - Suportă asocieri multiple (dacă aceeași piesă este folosită în mai multe ansambluri).

2. **Creează un folder dedicat pentru fiecare piesă**:
   - Numele folderului corespunde exact denumirii piesei (cu sanitizare automată pentru caractere nepermise de Windows, precum `/`, `*`, `:`, `?` etc.).

3. **Găsește și transferă fișierele corespunzătoare**:
   - **Piesa NC (`.nc`, `.nc1`)**: mutată/copiată direct în folderul piesei. Suportă caractere speciale și delimitatori (`-`, `_`).
   - **Desenele PDF ale ansamblurilor asociate**: toate desenele PDF ale ansamblurilor din care face parte piesa.
   - **Piesele DWG pe grosimi de tablă**:
     - Scanează subfolderele sursă ce reprezintă grosimi de tablă (`10`, `t10`, `PL12`, `15mm`, `thk 8`, `20` etc.).
     - Recunoaște piesa de bază și toate **variațiile sale CAD (T, L, C etc.)** (ex: `101`, `101T`, `101L`, `101C`, `101_T`, `101-L`, `101 C`), excluzând fals-pozitivele numerice (ex: `1015`).
     - Plasează DWG-urile în **subfoldere cu denumirea grosimii de tablă** în interiorul folderului piesei.

4. **Siguranță la Mutare (Zero Pierderi de Date)**:
   - În modul **Mutare (Move)**, desenele de ansamblu comune sunt multiplicate în siguranță pentru fiecare piesă, iar originalele din sursă sunt șterse abia la finalul procesului.

5. **Raport detaliat Excel**:
   - Generează automat `Raport_Organizare_Piese.xlsx` în destinație.

---

## 🔄 Modulul 2: Regrupare DWG pe Grosimi (Piese -> Grosimi)

**Funcția inversă:** luați folderul ce conține folderele de piese (ex: `Piese_Organizate`) și regrupați **DOAR fișierele DWG** pe grosimi de tablă:

1. **Scanare inteligentă**:
   - Parcurge recursiv toate folderele de piese.
   - Filtrează și extrage **exclusiv fișierele `.dwg`** (ignoră PDF-uri, fișiere NC și fișiere `.bak`).
2. **Detectare și Unificare Grosime**:
   - Identifică grosimea fiecărui DWG din subfolderul în care se află (ex. `T10`, `t10`, `12`, `PL15`, `T20` etc.).
   - **Unifică denumirile**: piesele din `t10` și cele din `T10` ajung în **același folder unificat `T10`**.
3. **Toate piesele de aceeași grosime în același folder**:
   - Toate piesele DWG cu grosimea respectivă (ex: 10mm) sunt puse împreună în folderul acelei grosimi.
4. **Protecție coliziuni de nume**:
   - Dacă două piese diferite au un desen cu același nume, fișierul este redenumit sigur (ex: `nume (Piesa).dwg`), astfel încât niciun desen să nu fie suprascris.
5. **Raport dedicat**:
   - Generează `Raport_Regrupare_DWG.xlsx` cu evidențierea fiecărui fișier, a grosimii și a piesei sursă din care provine.

---

## 📁 Structura înainte și după:

### Din Modulul 1 (Extragere pe Piese):
```text
Destinatie_Piese/
├── P-101/
│   ├── P-101.nc
│   ├── 1242 - 48.2.pdf
│   └── T10/
│       ├── P-101.dwg
│       ├── P-101T.dwg
│       └── P-101C.dwg
└── P-102/
    ├── P-102.nc1
    ├── 1242 - 48.2.pdf
    └── T15/
        └── P-102.dwg
```

### Din Modulul 2 (Regrupare Inversă DWG):
```text
Destinatie_Grosimi/
├── Raport_Regrupare_DWG.xlsx
├── T10/
│   ├── P-101.dwg
│   ├── P-101T.dwg
│   ├── P-101C.dwg
│   ├── DSB2004C.dwg
│   └── DSB2004T.dwg
├── T15/
│   ├── P-102.dwg
│   └── DSB2101T.dwg
└── T20/
    ├── DSB2014C1.dwg
    └── DSB2052C.dwg
```

---

## 🚀 Cum se rulează?

### Varianta 1: Executabilul direct (Fără instalare)
Faceți dublu click pe:
```text
MiniExtragereNC.exe
```

### Varianta 2: Fișierul Batch
Faceți dublu click pe:
```text
Start_Aplicatie.bat
```

### Varianta 3: Din terminal Python
```powershell
python app.py
```
