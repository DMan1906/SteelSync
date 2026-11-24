# NORSTAL — Automatizare Introducere Planuri de Croi

Aplicație locală Python cu interfață web **Streamlit** și automatizare **Playwright** pentru încărcarea, verificarea, validarea, inserarea codului QR în documentele Word Metalix și printarea automată a planurilor de croi în intranetul NORSTAL (`http://10.0.1.132/NORSTAL/`).

## Structura proiectului

- `app.py` — Interfața web dark-mode Streamlit (2 etape, preview editabil, log-uri în timp real, galerie QR).
- `settings_store.py` — Persistență locală automată în JSON (`user_settings.json` pentru setări/credențiale și `work_tmp/last_session.json` pentru lotul de planuri, editările din previzualizare și rezultatele Etapa 1/2).
- `parser_docs.py` — Conversie automată `.DOC` (Word 97-2003 OLE2 via Word COM) -> `.docx`, parsare `SubNest`, `Part weight`, `Machine`/`Laser` (`Plasma`/`Laser`), cod comandă și împerechere cu `.xlsx`.
- `browser_bot.py` — Clasă Playwright (`NorstalBrowserBot`) pe thread dedicat care păstrează sesiunea activă între Etapa 1 și Etapa 2, suportă `iframe_a`, Select2 și filtrare rânduri ROȘU vs ROZ.
- `word_qr_printer.py` — Inserează imaginea QR în celula `QR code` din tabelul `Laser parameters` (salvat în `output_print/`) și trimite opțional la imprimanta implicită din Windows.

## Instalare & Configurare

1. Creează și activează mediul virtual (dacă nu este deja creat):
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\pip.exe install -r requirements.txt
   .\.venv\Scripts\playwright.exe install chromium
   ```

2. Opțional, copiază `.env.example` în `.env` și completează credențialele intranetului:
   ```powershell
   Copy-Item .env.example .env
   ```

## Pornire aplicație

```powershell
.\.venv\Scripts\streamlit.exe run app.py
```

## Flux de lucru (2 Etape)

1. **Încărcare & Parsare**: Alege folderul local (ex: `d:\Programe plasma`) sau încarcă fișierele Word (`.doc`/`.docx`) și `.xlsx`. Verifică tabelul de previzualizare.
2. **Etapa 1**: Botul se autentifică în intranet, navighează în `Productie -> Introducere date -> Introducere planuri de croi`, creează planul (`Panou 1` sub debitare `Plasma`), adaugă comanda și greutatea (`Panou 2`), introduce piesele din `.xlsx` în `Panou 3` în funcție de opțiunea selectată în sidebar (**🚀 Full HTTP — tot planul P1+P2+P3 rapid** prin calibrare pe Planul #1 și replay HTTP pe #2..#N, **⚡ HTTP Request rapid (Panou 3)**, **📥 Import Excel (.xlsx)**, sau implicit **piesă cu piesă prin `+`** dacă niciuna nu este bifată) și șterge automat rândurile **ROȘII** (păstrând rândurile **ROZ**). Sesiunea de browser rămâne deschisă pentru verificare vizuală.
3. **Etapa 2**: După verificarea umană, apasă **„Validează, Inserează QR & Printează”**. Botul bifează `Valid` pe fiecare plan, extrage imaginea `QRcode`, o inserează în fișierul Word în `output_print/` și (dacă opțiunea de printare este bifată în sidebar) trimite documentele la imprimanta implicită Windows.

## Ajustare selectori DOM

Dacă intranetul suferă modificări de layout, toți selectorii sunt centralizați în dicționarul `SELECTORS` din partea de sus a fișierului `browser_bot.py`.
