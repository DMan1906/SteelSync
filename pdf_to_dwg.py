# -*- coding: utf-8 -*-
"""
pdf_to_dwg.py - Generator CAD (DWG / DXF) din Desene de Piesă PDF
Extrage doar geometria curată a piesei (1:1 în milimetri), fără cote, texte, săgeți de secțiune sau cartuș.
Suportă decupaje complexe, găuri, colțuri rotunjite (fillet arcs native), vederi scurtate Tekla cu 1 sau multiple linii de ruptură (multi-break)
și unirea continuă a muchiilor coliniare fără goluri sau linii suprapuse.
"""

import os
import re
import math
import glob
import shutil
from pathlib import Path
from collections import defaultdict

try:
    import fitz  # PyMuPDF
except ImportError:
    try:
        import pymupdf as fitz
    except ImportError:
        fitz = None

try:
    import ezdxf
    from ezdxf.addons import odafc
except ImportError:
    ezdxf = None
    odafc = None


class PdfToCadConverter:
    """Motor de conversie din desen PDF de piesă în fișier CAD (DWG / DXF) 1:1."""

    # Cuvinte cheie / denumiri excluse categoric din a fi considerate coduri de reper
    BLACKLIST_NAMES = {
        "POS", "NO", "ANTALL", "PROFIL", "PROFILE", "MATERIAL", "LENGDE", "LENGTH",
        "LANGD", "LÄNGD", "VEKT", "WEIGHT", "VIKT", "TOTALVIKT", "AREA", "LIT", "LITTRA",
        "DEL", "STORLEK", "TOTALANTAL", "NUMBER", "ASSEMBLY", "TOASSEMBLY",
        "NORSKE", "BARESYSTEMER", "BÆRESYSTEMER", "SWECO", "FISKEBRYGGA", "STALPLATE",
        "STÅLPLATE", "RIB", "A4", "A3", "A2", "A1", "A0", "FORMAT", "DATO", "DATUM",
        "REV", "ENDRING", "UTF", "KONTR", "STATUS", "ANSVARLIG", "TITTEL", "OPPDRAGSGIVER",
        "OPPDRAGSNR", "UPPDRAG", "UPPDRAGSNR", "UPPDRAGSNUMMER", "TEGNINGSNUMMER",
        "TEGNINGSSTATUS", "ORIGINALFORMAT", "ORGINALFORMAT", "ARBEIDSTEGNING",
        "SKALA", "HANDLÄGGARE", "HANDLAGGARE", "LITTERA", "BYGGHANDLING", "BET", "SIGN",
        "RITAD", "NUMMER", "SPEC", "REVISION", "DRAWINGNUMBER", "PROJECTNUMBER"
    }

    def __init__(self, oda_path=None):
        self.oda_path = oda_path
        self._configure_odafc()

    @classmethod
    def _find_default_oda_exe(cls):
        """Caută automat executabilul ODAFileConverter în locațiile standard de instalare Windows."""
        for pattern in [
            r"C:\Program Files\ODA\*\ODAFileConverter.exe",
            r"C:\Program Files (x86)\ODA\*\ODAFileConverter.exe",
        ]:
            matches = glob.glob(pattern)
            for m in matches:
                if os.path.isfile(m):
                    return m
        return None

    def _configure_odafc(self):
        """Configurează calea către ODA File Converter dacă este disponibilă."""
        if odafc is None or ezdxf is None:
            return
        target_path = self.oda_path
        if not target_path or not os.path.isfile(target_path):
            target_path = self._find_default_oda_exe()
        if target_path and os.path.isfile(target_path):
            try:
                self.oda_path = target_path
                ezdxf.options.set("odafc-addon", "win_exec_path", target_path)
                odafc.win_exec_path = target_path
            except Exception:
                pass

    @staticmethod
    def is_available():
        """Verifică dacă bibliotecile necesare (PyMuPDF și ezdxf) sunt instalate."""
        return fitz is not None and ezdxf is not None

    def has_oda(self):
        return self.is_odafc_available(self.oda_path)

    @classmethod
    def is_odafc_available(cls, custom_path=None):
        """Verifică dacă ODA File Converter este configurat și funcțional pentru DWG nativ."""
        if odafc is None or ezdxf is None:
            return False
        target_path = custom_path
        if not target_path or not os.path.isfile(target_path):
            target_path = cls._find_default_oda_exe()
        if target_path and os.path.isfile(target_path):
            try:
                ezdxf.options.set("odafc-addon", "win_exec_path", target_path)
                odafc.win_exec_path = target_path
            except Exception:
                pass
        try:
            return odafc.is_installed()
        except Exception:
            return False

    @classmethod
    def clean_part_name(cls, raw_name):
        """Curăță denumirea piesei de cratime, spații și caractere speciale (ex: DSBO-21051 -> DSBO21051)."""
        if not raw_name:
            return "PART"
        clean = re.sub(r'[^A-Za-z0-9]', '', raw_name).upper()
        if clean in cls.BLACKLIST_NAMES:
            return "PART"
        return clean if clean else "PART"

    def extract_drawing_metadata(self, page, pdf_path):
        """
        Extrage metadatele din textul paginii:
        - Codul piesei (POS. NO. / POS. NUMMER / Nume fișier)
        - Scara desenului (Målestokk 1:X, Front 1:X)
        - Dimensiunile nominale (LENGDE, PROFIL: grosime, lățime)
        - Coordonata Y de început a tabelului STYKKLISTE
        """
        text = page.get_text()
        blocks = page.get_text("blocks")
        stem = Path(pdf_path).stem
        clean_stem = re.sub(r'[^A-Za-z0-9]', '', stem).upper()

        meta = {
            "pos_no": None,
            "scale": 10.0,
            "lengde": None,
            "profil_width": None,
            "profil_thickness": None,
            "stykk_y": None,
            "top_margin_y": 20.0,
            "clean_name": None,
        }

        # 1. Determinare Y pentru STYKKLISTE / Cartuș (diferențiere sus vs jos)
        for b in blocks:
            txt = b[4].lower()
            if b[1] < page.rect.height * 0.35 and any(
                k in txt for k in ["stykkliste", "storlek", "totalantal", "total vikt", "längd (mm)", "langd (mm)"]
            ):
                meta["top_margin_y"] = max(meta["top_margin_y"], b[3] + 8.0)
            elif b[1] >= page.rect.height * 0.40 and any(
                k in txt for k in [
                    "stykkliste", "arbeidstegning", "pos.nummer", "originalformat", "orginalformat",
                    "ändringen avser", "bygghandling", "handläggare", "ritad av", "uppdrag nr",
                    "uppdragsnummer", "revisjon", "norconsult", "no/ pcs"
                ]
            ):
                if meta["stykk_y"] is None or b[1] < meta["stykk_y"]:
                    meta["stykk_y"] = b[1]

        # 1.5 Aliniere spațială exactă pe coloana LENGDE / LÄNGD / LENGTH din tabelul BOM
        words = page.get_text("words")
        lengde_headers = [w for w in words if re.match(r'^(?:LENGDE|LÄNGD|LANGD|LENGTH)$', w[4], re.IGNORECASE)]
        if not lengde_headers:
            lengde_headers = [w for w in words if any(k in w[4].upper() for k in ("LENGDE", "LÄNGD", "LANGD"))]

        spatial_candidates = []
        for hw in lengde_headers:
            hx0, hy0, hx1, hy1, htxt = hw[:5]
            for w in words:
                wx0, wy0, wx1, wy1, wtxt = w[:5]
                # Numărul de sub antetul LENGDE în interval de 45 pt pe verticală
                if hy1 < wy0 < hy1 + 45:
                    dist_x = abs((wx0 + wx1) / 2.0 - (hx0 + hx1) / 2.0)
                    clean_txt = wtxt.replace(',', '.')
                    if clean_txt.replace('.', '', 1).isdigit():
                        val = float(clean_txt)
                        if val > 15:
                            spatial_candidates.append((dist_x, val))
        if spatial_candidates:
            spatial_candidates.sort(key=lambda c: c[0])
            meta["lengde"] = spatial_candidates[0][1]

        # 2. Extragere din blocul de date STYKKLISTE
        for b in blocks:
            b_txt = b[4].strip()
            lines = [l.strip() for l in b_txt.splitlines() if l.strip()]
            for idx, line in enumerate(lines):
                m_prof = re.search(r'\b(?:PL|BL|FL|P|FLAT)\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\b', line, re.IGNORECASE)
                if not m_prof:
                    m_prof = re.search(r'\b(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\b', line, re.IGNORECASE)
                if m_prof:
                    try:
                        d1 = float(m_prof.group(1))
                        d2 = float(m_prof.group(2))
                        meta["profil_thickness"] = min(d1, d2)
                        meta["profil_width"] = max(d1, d2)
                    except ValueError:
                        pass

                    if idx > 0:
                        # Căutăm în sus primul rând care nu este un simplu număr/greutate (ex. 6.6)
                        for look_back in range(idx - 1, -1, -1):
                            prev = lines[look_back].strip()
                            clean_prev = re.sub(r'[^A-Za-z0-9]', '', prev).upper()
                            if clean_prev.replace('.', '').isdigit():
                                continue
                            if clean_prev and clean_prev not in self.BLACKLIST_NAMES and any(c.isdigit() for c in clean_prev):
                                meta["pos_no"] = prev
                                break

                    if not meta["lengde"]:
                        candidates_sub = []
                        for sub in lines[idx + 1:]:
                            clean_sub = sub.replace(',', '.')
                            if clean_sub.replace('.', '', 1).isdigit():
                                val = float(clean_sub)
                                if val >= 50:
                                    candidates_sub.append(val)
                        if candidates_sub:
                            # Lungimea reală este numărul mare din rând (nu cantitatea/antall)
                            meta["lengde"] = max(candidates_sub)
                    break
            if meta["pos_no"]:
                break

        # 3. Căutare după Tegningsnummer în cartuș dacă pos_no nu a fost găsit încă
        if not meta["pos_no"]:
            m_teg = re.search(r'(?:Pos\.\s*nummer|Tegn\.nr\.?|Tegningsnummer)\s*[\r\n\:]+\s*([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
            if m_teg:
                candidate = m_teg.group(1).strip()
                if self.clean_part_name(candidate) not in self.BLACKLIST_NAMES:
                    meta["pos_no"] = candidate

        # 4. Căutare 'STYKKLISTE FOR <COD>'
        if not meta["pos_no"]:
            m_for = re.search(r'STYKKLISTE\s+FOR\s+([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
            if m_for:
                candidate = m_for.group(1).strip()
                if self.clean_part_name(candidate) not in self.BLACKLIST_NAMES:
                    meta["pos_no"] = candidate

        # 5. Prioritate / Fallback Nume Fișier PDF
        if clean_stem and clean_stem not in self.BLACKLIST_NAMES and any(c.isdigit() for c in clean_stem):
            if not meta["pos_no"] or meta["pos_no"].upper() in self.BLACKLIST_NAMES:
                meta["pos_no"] = clean_stem
        elif not meta["pos_no"]:
            meta["pos_no"] = stem

        meta["clean_name"] = self.clean_part_name(meta["pos_no"])

        # 6. Identificare Scară (căutare spațială sub SKALA / SCALE / MÅLESTOKK + regex text)
        found_scale = None
        scale_headers = [w for w in words if re.match(r'^(?:SKALA|SCALE|MÅLESTOKK|MALESTOKK):?$', w[4], re.IGNORECASE)]
        scale_candidates = []
        for hw in scale_headers:
            hx0, hy0, hx1, hy1, _ = hw[:5]
            for w in words:
                wx0, wy0, wx1, wy1, wtxt = w[:5]
                m_w = re.match(r'^1\s*:\s*(\d+(?:\.\d+)?)$', wtxt.strip())
                if m_w and (hy0 - 5.0 <= wy0 <= hy1 + 35.0) and abs(wx0 - hx0) <= 105.0:
                    try:
                        val = float(m_w.group(1))
                        if 0.1 <= val <= 100.0:
                            dist = abs(wx0 - hx0) + abs(wy0 - hy0)
                            scale_candidates.append((dist, val))
                    except ValueError:
                        pass
        if scale_candidates:
            scale_candidates.sort(key=lambda c: c[0])
            found_scale = scale_candidates[0][1]

        if found_scale is None:
            m_scale = re.search(r'(?:M.*?lestokk|Skala|Scale|Front|Top|View)\s*[:\s]?\s*1\s*:\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
            if m_scale:
                try:
                    found_scale = float(m_scale.group(1))
                except ValueError:
                    found_scale = None

        # Fallback: etichetă de scară solitară 1:X în zona de desen (deasupra cartușului)
        if found_scale is None:
            limit_y = meta["stykk_y"] if meta["stykk_y"] else page.rect.height * 0.75
            for w in words:
                if w[1] < limit_y:
                    m_w = re.match(r'^1:(\d+(?:\.\d+)?)$', w[4].strip())
                    if m_w:
                        try:
                            val = float(m_w.group(1))
                            if 1.0 <= val <= 50.0:
                                found_scale = val
                                break
                        except ValueError:
                            pass

        meta["scale"] = found_scale if found_scale is not None else 10.0

        if not meta["lengde"]:
            m_len = re.search(r'(?:LENGDE|LÄNGD|LANGD|LENGTH)\s*[\r\n]+\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
            if m_len:
                try:
                    meta["lengde"] = float(m_len.group(1))
                except ValueError:
                    pass

        if not meta["profil_width"]:
            m_prof = re.search(r'\b(?:PL|BL|FL|P|FLAT)\s*(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\b', text, re.IGNORECASE)
            if not m_prof:
                m_prof = re.search(r'\b(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)\b', text, re.IGNORECASE)
            if m_prof:
                try:
                    d1 = float(m_prof.group(1))
                    d2 = float(m_prof.group(2))
                    meta["profil_thickness"] = min(d1, d2)
                    meta["profil_width"] = max(d1, d2)
                except ValueError:
                    pass

        return meta

    @staticmethod
    def _is_hole_circle_drawing(d):
        """Returnează (cx, cy, radius) dacă desenul reprezintă o gaură circulară (4 sau 6 curbe Bezier), altfel None."""
        items = d.get("items", [])
        n_curves = sum(1 for it in items if it[0] == "c")
        if (len(items) == 4 and n_curves == 4) or (n_curves == 4 and len(items) <= 6) or (len(items) == 6 and n_curves == 6):
            r = d["rect"]
            rx = (r.x1 - r.x0) / 2.0
            ry = (r.y1 - r.y0) / 2.0
            min_r = 0.8 if len(items) > 4 else 0.4
            if rx > 0 and ry > 0 and abs(rx - ry) < 0.35 * max(rx, ry) and min_r <= rx <= 45.0:
                rad = min(rx, ry) if n_curves == 6 else (rx + ry) / 2.0
                return ((r.x0 + r.x1) / 2.0, (r.y0 + r.y1) / 2.0, rad)
        return None

    @staticmethod
    def _is_crosshair_line(d, holes_list):
        """Detectează liniile de axă / centrare (+) care trec prin centrul unei găuri sau coboară la cotă."""
        items = d.get("items", [])
        if len(items) != 1 or items[0][0] != "l" or not holes_list:
            return False
        p1, p2 = items[0][1], items[0][2]
        vx, vy = p2.x - p1.x, p2.y - p1.y
        seg_len = math.hypot(vx, vy)
        if seg_len < 1.0:
            return False
        ux, uy = vx / seg_len, vy / seg_len
        for cx, cy, rad in holes_list:
            if seg_len <= max(240.0, 6.0 * rad):
                dx, dy = cx - p1.x, cy - p1.y
                proj = dx * ux + dy * uy
                if 0.25 * rad <= proj <= seg_len - 0.25 * rad:
                    perp = abs(dx * uy - dy * ux)
                    if perp <= 1.2:
                        return True
            if abs(vx) < 0.5 and abs(p1.x - cx) <= 0.8 and seg_len <= 240.0:
                y_top = min(p1.y, p2.y)
                y_bot = max(p1.y, p2.y)
                if y_top <= cy + 2.0 * rad and y_top >= cy - 2.5 * rad and y_bot > cy + rad:
                    return True
        return False

    @classmethod
    def _detect_frame_and_table_exclusions(cls, page, drawings):
        """
        Detectează chenarul interior al foii și dreptunghiurile tabelelor/cartușelor (suedeză/norvegiană/engleză)
        pentru a elimina liniuțele de centrare exterioare și liniile de tabel (ex. TO ASSEMBLY, ÄNDRINGEN AVSER).
        """
        pw, ph = page.rect.width, page.rect.height
        left_x, right_x, top_y, bot_y = [], [], [], []
        hv_lines = []
        frame_border_indices = set()

        for idx, d in enumerate(drawings):
            items = d.get("items", [])
            if len(items) == 1 and items[0][0] == "l":
                p1, p2 = items[0][1], items[0][2]
                length = math.hypot(p2.x - p1.x, p2.y - p1.y)
                is_h = abs(p1.y - p2.y) < 0.5
                is_v = abs(p1.x - p2.x) < 0.5
                ym = (p1.y + p2.y) / 2.0
                xm = (p1.x + p2.x) / 2.0
                is_frame_line = (
                    (is_h and length > 0.80 * pw and (ym < ph * 0.14 or ym > ph * 0.84)) or
                    (is_v and length > 0.80 * ph and (xm < pw * 0.14 or xm > pw * 0.84))
                )
                if is_frame_line:
                    frame_border_indices.add(idx)
                    if is_h:
                        if 2.0 < ym < ph * 0.14:
                            top_y.append(ym)
                        elif ph * 0.84 < ym < ph - 2.0:
                            bot_y.append(ym)
                    else:
                        if 2.0 < xm < pw * 0.14:
                            left_x.append(xm)
                        elif pw * 0.84 < xm < pw - 2.0:
                            right_x.append(xm)
                elif is_h or is_v:
                    hv_lines.append((idx, p1, p2, is_h, is_v, length))

        frame = (
            max(left_x) if left_x else None,
            max(top_y) if top_y else None,
            min(right_x) if right_x else None,
            min(bot_y) if bot_y else None,
        )

        # Identificare blocuri de text ce aparțin tabelelor / cartușelor
        table_kw = [
            "to assembly", "number\nto assembly", "ändringen avser", "bygghandling",
            "handläggare", "ritad av", "uppdrag nr", "uppdragsnummer", "total vikt",
            "storlek", "totalantal", "stykkliste", "arbeidstegning", "pos.nummer",
            "pos. nr", "originalformat", "orginalformat", "norconsult", "revisjon",
            "drawingnumber", "projectnumber", "längd (mm)", "langd (mm)"
        ]
        table_blocks = []
        to_asm_blocks = []
        try:
            for b in page.get_text("blocks"):
                bt = b[4].lower()
                if any(k in bt for k in table_kw):
                    table_blocks.append((b[0], b[1], b[2], b[3]))
                    if "to assembly" in bt:
                        to_asm_blocks.append((b[0], b[1], b[2], b[3]))
        except Exception:
            pass

        # Grupare linii H/V (fără liniile lungi de chenar) după conexiuni de capăt
        n_hv = len(hv_lines)
        parent = list(range(n_hv))

        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i

        def union(i, j):
            ri, rj = find(i), find(j)
            if ri != rj:
                parent[rj] = ri

        grid = defaultdict(list)
        cell = 1.0
        for hi, (_, p1, p2, _, _, _) in enumerate(hv_lines):
            for pt in (p1, p2):
                kx, ky = int(math.floor(pt.x / cell)), int(math.floor(pt.y / cell))
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        for oj in grid.get((kx + dx, ky + dy), []):
                            union(hi, oj)
                grid[(kx, ky)].append(hi)

        comps = defaultdict(list)
        for hi in range(n_hv):
            comps[find(hi)].append(hv_lines[hi])

        excluded_indices = set(frame_border_indices)
        fx0, fy0, fx1, fy1 = frame
        table_rects = []

        for bx0, by0, bx1, by1 in to_asm_blocks:
            right_edge = (fx1 + 2.0) if fx1 is not None else pw
            table_rects.append((bx0 - 15.0, by0 - 12.0, right_edge, min(by1 + 65.0, ph * 0.26)))

        for comp in comps.values():
            xs = [pt.x for _, p1, p2, _, _, _ in comp for pt in (p1, p2)]
            ys = [pt.y for _, p1, p2, _, _, _ in comp for pt in (p1, p2)]
            cx0, cx1, cy0, cy1 = min(xs), max(xs), min(ys), max(ys)
            comp_w = cx1 - cx0
            comp_h = cy1 - cy0

            if comp_w > 0.70 * pw and comp_h > 0.70 * ph:
                for d_idx, _, _, _, _, _ in comp:
                    excluded_indices.add(d_idx)
                continue

            is_table = False
            # Componentă care conține un text de antet tabel / cartuș
            for bx0, by0, bx1, by1 in table_blocks:
                if cx0 - 8.0 <= bx0 and bx1 <= cx1 + 8.0 and cy0 - 8.0 <= by0 and by1 <= cy1 + 8.0:
                    is_table = True
                    break

            if is_table:
                for d_idx, _, _, _, _, _ in comp:
                    excluded_indices.add(d_idx)
                if comp_h < 0.45 * ph:
                    table_rects.append((cx0 - 1.5, cy0 - 1.5, cx1 + 1.5, cy1 + 1.5))

        # Excludem orice altă linie interioară dintr-un dreptunghi de tabel identificat
        if table_rects:
            for idx, d in enumerate(drawings):
                if idx in excluded_indices:
                    continue
                r = d["rect"]
                for tx0, ty0, tx1, ty1 in table_rects:
                    if tx0 <= r.x0 and r.x1 <= tx1 and ty0 <= r.y0 and r.y1 <= ty1:
                        excluded_indices.add(idx)
                        break

        return frame, excluded_indices

    def extract_geometry_items(self, page, stykk_y=None, top_margin_y=20.0, only_front_view=True, meta=None):
        """
        Extrage geometria curată a piesei:
        - Filtrează chenarul paginii, marcajele de centrare, tabelele (BOM, TO ASSEMBLY, revizii) și cartușul.
        - Elimină automat liniile de cotă, axele de centrare ale găurilor (+), săgețile și simbolurile de secțiune.
        - Izolează vederea principală (Front) față de secțiunile laterale înguste și vederea Topp.
        - Recuperează arcele de racordare (fillet) desenate cu linie subțire atașate conturului.
        """
        max_y = (stykk_y - 6) if stykk_y else (page.rect.height * 0.88)
        margin_x0 = 20.0
        margin_x1 = page.rect.width - 20.0
        margin_y0 = top_margin_y if top_margin_y else 20.0

        drawings = page.get_drawings()
        frame, table_excluded = self._detect_frame_and_table_exclusions(page, drawings)
        fx0, fy0, fx1, fy1 = frame
        if fx0 is not None:
            margin_x0 = max(margin_x0, fx0 + 1.0)
        if fx1 is not None:
            margin_x1 = min(margin_x1, fx1 - 1.0)
        if fy0 is not None:
            margin_y0 = max(margin_y0, fy0 + 1.0)
        if fy1 is not None:
            max_y = min(max_y, fy1 - 1.0)

        graphic_drawings = []
        holes_list = []
        hole_drawing_indices = set()
        for idx, d in enumerate(drawings):
            if idx in table_excluded:
                continue
            r = d["rect"]
            if not (r.y1 < max_y and r.y0 > margin_y0 and r.x0 > margin_x0 and r.x1 < margin_x1):
                continue
            # Excludem eventuale linii lungi de chenar rămase
            if (r.x1 - r.x0) > page.rect.width * 0.70 or (r.y1 - r.y0) > page.rect.height * 0.70:
                continue
            hc = self._is_hole_circle_drawing(d)
            if hc is not None:
                holes_list.append(hc)
                hole_drawing_indices.add(idx)
            graphic_drawings.append((idx, d))

        if not graphic_drawings:
            return [], None

        # 1. Eliminare simboluri groase de secțiune (> 1.30 pt), cercuri de gaură, nori de revizie și axe de centru (+)
        non_section_drawings = []
        for idx, d in graphic_drawings:
            if idx in hole_drawing_indices:
                continue
            w = d.get("width") or 0
            if w <= 0 or w > 1.30:
                continue
            if sum(1 for it in d.get("items", []) if it[0] == "c") >= 8:
                continue
            if self._is_crosshair_line(d, holes_list):
                continue
            non_section_drawings.append(d)

        if not non_section_drawings:
            non_section_drawings = [
                d for idx, d in graphic_drawings
                if idx not in hole_drawing_indices and (d.get("width") or 0) > 0
            ]

        if not non_section_drawings:
            return [], None

        # 2. Grupare după grosime (clustere +/- 0.05 pt) și alegerea grosimii reale de contur
        raw_widths = sorted(set(round(d.get("width") or 0, 3) for d in non_section_drawings if d.get("width")))
        width_bands = []
        for rw in raw_widths:
            if rw < 0.40:
                continue
            if not width_bands or rw - width_bands[-1][1] > 0.05:
                width_bands.append([rw, rw])
            else:
                width_bands[-1][1] = rw

        def _score_width_band(w_min, w_max):
            drs = [
                d for d in non_section_drawings
                if d.get("width") and (w_min - 0.01) <= round(d.get("width"), 3) <= (w_max + 0.01) and d.get("items")
            ]
            if not drs:
                return -1e9, []
            xs = [d["rect"].x0 for d in drs] + [d["rect"].x1 for d in drs]
            ys = [d["rect"].y0 for d in drs] + [d["rect"].y1 for d in drs]
            span_x = max(xs) - min(xs)
            span_y = max(ys) - min(ys)
            if span_x < 6.0 or span_y < 6.0:
                return -1e8, drs

            deg = defaultdict(int)
            seen_segs = set()
            h_right_end = set()
            h_left_end = set()
            n_items = 0
            has_h_extent = False
            has_v_extent = False
            for d in drs:
                for it in d["items"]:
                    if it[0] in ("l", "c"):
                        p_start = it[1]
                        p_end = it[2] if it[0] == "l" else it[4]
                        ks = (round(p_start.x * 2.0) / 2.0, round(p_start.y * 2.0) / 2.0)
                        ke = (round(p_end.x * 2.0) / 2.0, round(p_end.y * 2.0) / 2.0)
                        if ks == ke and it[0] == "l":
                            continue
                        seg_key = (it[0], min(ks, ke), max(ks, ke))
                        if seg_key in seen_segs:
                            continue
                        seen_segs.add(seg_key)
                        n_items += 1
                        if abs(p_end.x - p_start.x) > 2.0:
                            has_h_extent = True
                        if abs(p_end.y - p_start.y) > 2.0:
                            has_v_extent = True
                        if it[0] == "l" and abs(p_end.y - p_start.y) < 0.5 and abs(p_end.x - p_start.x) > 2.0:
                            if ks[0] <= ke[0]:
                                h_left_end.add(ks)
                                h_right_end.add(ke)
                            else:
                                h_left_end.add(ke)
                                h_right_end.add(ks)
                        deg[ks] += 1
                        deg[ke] += 1

            if n_items < 3 or not (has_h_extent and has_v_extent):
                return -1e8, drs

            clean_verts = sum(1 for v in deg.values() if v == 2)
            # Capetele de grad 1 care formează perechi de ruptură Tekla pe orizontală sunt noduri valide
            deg1_pts = sorted([pt for pt, v in deg.items() if v == 1])
            used_d1 = set()
            for idx_a in range(len(deg1_pts)):
                if idx_a in used_d1:
                    continue
                pt_a = deg1_pts[idx_a]
                if pt_a not in h_right_end:
                    continue
                ax, ay = pt_a
                for idx_b in range(idx_a + 1, len(deg1_pts)):
                    if idx_b in used_d1:
                        continue
                    pt_b = deg1_pts[idx_b]
                    if pt_b not in h_left_end:
                        continue
                    bx, by = pt_b
                    if abs(ay - by) <= 0.6 and 1.2 <= (bx - ax) <= 16.0:
                        clean_verts += 2
                        used_d1.add(idx_a)
                        used_d1.add(idx_b)
                        break

            total_verts = max(1, len(deg))
            clean_ratio = clean_verts / float(total_verts)
            w_mid = (w_min + w_max) / 2.0
            score = (
                clean_ratio * 1000.0
                + (200.0 if w_mid >= 0.54 else 0.0)
                + min(span_x, 600.0) * 0.8
                + w_mid * 50.0
                - max(0, sum(1 for v in deg.values() if v >= 3)) * 25.0
            )
            return score, drs

        contour_drawings = []
        if width_bands:
            scored_bands = [(_score_width_band(b[0], b[1]), b) for b in width_bands]
            scored_bands.sort(key=lambda x: x[0][0], reverse=True)
            best_score, contour_drawings = scored_bands[0][0]
            target_w_min = scored_bands[0][1][0] - 0.01
        else:
            target_w_min = 0.40

        if not contour_drawings:
            contour_drawings = [d for d in non_section_drawings if d.get("width") and d.get("items")]

        # 3. Izolare vedere principală (Front View) prin clustering orizontal pe axa X
        if only_front_view and len(contour_drawings) > 1:
            intervals = sorted([(d["rect"].x0, d["rect"].x1, d) for d in contour_drawings], key=lambda x: x[0])
            clusters = []
            cur_x0, cur_x1 = intervals[0][0], intervals[0][1]
            cur_drawings = [intervals[0][2]]
            for x0, x1, dr in intervals[1:]:
                if x0 - cur_x1 <= 25.0:  # aparțin aceleiași vederi
                    cur_x1 = max(cur_x1, x1)
                    cur_drawings.append(dr)
                else:
                    clusters.append((cur_x0, cur_x1, cur_drawings))
                    cur_x0, cur_x1 = x0, x1
                    cur_drawings = [dr]
            clusters.append((cur_x0, cur_x1, cur_drawings))

            # Selectăm clusterul cel mai lat (Front view)
            clusters.sort(key=lambda c: (c[1] - c[0]), reverse=True)
            contour_drawings = clusters[0][2]

        # 3.1 Izolare vedere principală pe axa Y (eliminare vederi secundare cant/grosime t, Topp/Bunn sau secțiuni înguste A-A)
        if only_front_view and len(contour_drawings) > 1:
            intervals_y = sorted([(d["rect"].y0, d["rect"].y1, d) for d in contour_drawings], key=lambda y: y[0])
            clusters_y = []
            cur_y0, cur_y1 = intervals_y[0][0], intervals_y[0][1]
            cur_drawings_y = [intervals_y[0][2]]
            for y0, y1, dr in intervals_y[1:]:
                if y0 - cur_y1 <= 25.0:  # aparțin aceleiași vederi pe Y
                    cur_y1 = max(cur_y1, y1)
                    cur_drawings_y.append(dr)
                else:
                    clusters_y.append((cur_y0, cur_y1, cur_drawings_y))
                    cur_y0, cur_y1 = y0, y1
                    cur_drawings_y = [dr]
            clusters_y.append((cur_y0, cur_y1, cur_drawings_y))

            def _cluster_x_span(c):
                dr_list = c[2]
                return max(d["rect"].x1 for d in dr_list) - min(d["rect"].x0 for d in dr_list)

            max_span = max(_cluster_x_span(c) for c in clusters_y) if clusters_y else 1.0

            front_labels = []
            try:
                for b in page.get_text("blocks"):
                    if re.search(r"\bFront\b", b[4], re.IGNORECASE):
                        front_labels.append({"x0": b[0], "y0": b[1], "x1": b[2], "y1": b[3]})
            except Exception:
                pass

            scale_val = (meta.get("scale", 10.0) if meta else 10.0) or 10.0
            base_scale = scale_val * (25.4 / 72.0)
            target_pw = meta.get("profil_width") if meta else None
            target_th = meta.get("profil_thickness") if meta else None

            def _view_score(c):
                span = _cluster_x_span(c)
                if max_span > 0 and span < 0.65 * max_span:
                    return -100000.0 + span

                score = span
                h_pt = max(0.1, c[1] - c[0])
                h_mm = h_pt * base_scale

                if target_pw and target_pw > 0:
                    if abs(h_mm - target_pw) <= max(5.0, target_pw * 0.15):
                        score += 50000.0
                if target_th and target_th > 0 and (not target_pw or target_pw > target_th * 1.5):
                    if abs(h_mm - target_th) <= max(2.0, target_th * 0.2):
                        score -= 25000.0

                c_min_x = min(d["rect"].x0 for d in c[2])
                c_max_x = max(d["rect"].x1 for d in c[2])
                c_y0, c_y1 = c[0], c[1]

                for fl in front_labels:
                    if fl["x0"] <= c_max_x + 50.0 and fl["x1"] >= c_min_x - 50.0:
                        if 0 <= fl["y0"] - c_y1 <= 90.0:
                            score += 20000.0
                        elif 0 <= c_y0 - fl["y1"] <= 60.0:
                            score += 10000.0

                score += h_pt * 2.0
                return score

            clusters_y.sort(key=_view_score, reverse=True)
            contour_drawings = clusters_y[0][2]

        # 4. Colectare primitive geometrice
        items = []
        for d in contour_drawings:
            for it in d.get("items", []):
                items.append(it)

        if not items:
            return [], None

        # 5. Calcul Bounding Box
        pts = []
        for it in items:
            cmd = it[0]
            if cmd == "l":
                pts.extend([it[1], it[2]])
            elif cmd == "c":
                pts.extend([it[1], it[2], it[3], it[4]])
            elif cmd == "re":
                r = it[1]
                pts.extend([fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y1)])

        if not pts:
            return [], None

        bbox = {
            "min_x": min(p.x for p in pts),
            "max_x": max(p.x for p in pts),
            "min_y": min(p.y for p in pts),
            "max_y": max(p.y for p in pts),
        }

        # 5.1 Recuperare arce de racordare (fillet) desenate cu linie subțire (w < target_w_min)
        # care unesc două capete deschise (grad 1) ale conturului (ex. R10 în HSQP.1002)
        end_counts = defaultdict(list)
        for it in items:
            if it[0] in ("l", "c"):
                p_s = it[1]
                p_e = it[2] if it[0] == "l" else it[4]
                ks = (round(p_s.x, 1), round(p_s.y, 1))
                ke = (round(p_e.x, 1), round(p_e.y, 1))
                end_counts[ks].append(p_s)
                end_counts[ke].append(p_e)

        open_pts = [pts_list[0] for pts_list in end_counts.values() if len(pts_list) == 1]
        if len(open_pts) >= 2:
            thin_items = []
            for d in non_section_drawings:
                w_d = d.get("width") or 0
                if w_d >= target_w_min:
                    continue
                r = d["rect"]
                if not (bbox["min_x"] - 3.0 <= r.x0 and r.x1 <= bbox["max_x"] + 3.0 and
                        bbox["min_y"] - 3.0 <= r.y0 and r.y1 <= bbox["max_y"] + 3.0):
                    continue
                for it in d.get("items", []):
                    if it[0] == "c":
                        thin_items.append((it, it[1], it[4]))
                    elif it[0] == "l" and math.hypot(it[2].x - it[1].x, it[2].y - it[1].y) <= 8.0:
                        thin_items.append((it, it[1], it[2]))

            if thin_items:
                used_thin = set()
                recovered = []
                for idx_s in range(len(open_pts)):
                    p_src = open_pts[idx_s]
                    # BFS până la max 6 segmente pentru a atinge un alt capăt deschis
                    queue = [(p_src, [], set())]
                    found_path = None
                    while queue and not found_path:
                        curr_pt, path_idxs, visited_local = queue.pop(0)
                        if len(path_idxs) > 0:
                            for idx_t in range(len(open_pts)):
                                if idx_t == idx_s:
                                    continue
                                if math.hypot(curr_pt.x - open_pts[idx_t].x, curr_pt.y - open_pts[idx_t].y) <= 0.35:
                                    if any(thin_items[ti][0][0] == "c" for ti in path_idxs):
                                        found_path = path_idxs
                                        break
                        if found_path or len(path_idxs) >= 6:
                            continue
                        for ti, (it, ts, te) in enumerate(thin_items):
                            if ti in used_thin or ti in visited_local:
                                continue
                            if math.hypot(curr_pt.x - ts.x, curr_pt.y - ts.y) <= 0.25:
                                queue.append((te, path_idxs + [ti], visited_local | {ti}))
                            elif math.hypot(curr_pt.x - te.x, curr_pt.y - te.y) <= 0.25:
                                queue.append((ts, path_idxs + [ti], visited_local | {ti}))
                    if found_path:
                        for ti in found_path:
                            used_thin.add(ti)
                            recovered.append(thin_items[ti][0])

                if recovered:
                    items.extend(recovered)
                    for it in recovered:
                        if it[0] == "l":
                            pts.extend([it[1], it[2]])
                        elif it[0] == "c":
                            pts.extend([it[1], it[2], it[3], it[4]])
                    bbox = {
                        "min_x": min(p.x for p in pts),
                        "max_x": max(p.x for p in pts),
                        "min_y": min(p.y for p in pts),
                        "max_y": max(p.y for p in pts),
                    }

        bbox["width"] = bbox["max_x"] - bbox["min_x"]
        bbox["height"] = bbox["max_y"] - bbox["min_y"]

        return items, bbox

    @staticmethod
    def _find_view_break_stations(items, bbox):
        """
        Detectează dacă vederea este scurtată de Tekla (view shortening break)
        și localizează stațiile de ruptură (break stations) de-a lungul axei X.
        O ruptură Tekla are un spațiu între 1.2 și 16.0 pt (~0.4 - 5.6 mm pe hârtie, acoperind șabloane SUEDIA și NORVEGIA)
        și NU are linii de contur verticale/curbe conectate la capete.
        Returnează: listă de stații [{'xl': ..., 'xr': ..., 'center': ..., 'pairs': [...]}]
        """
        h_lines = []
        v_lines = []
        for idx, it in enumerate(items):
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.25:
                    h_lines.append((idx, min(p1.x, p2.x), max(p1.x, p2.x), (p1.y + p2.y) / 2.0, p1, p2))
                elif abs(p1.x - p2.x) < 0.25:
                    v_lines.append((idx, min(p1.y, p2.y), max(p1.y, p2.y), (p1.x + p2.x) / 2.0, p1, p2))

        break_pairs = []
        for i, (idx1, x1_min, x1_max, y1, p1a, p1b) in enumerate(h_lines):
            for j, (idx2, x2_min, x2_max, y2, p2a, p2b) in enumerate(h_lines):
                if idx1 == idx2:
                    continue
                if abs(y1 - y2) < 0.3:
                    gap = x2_min - x1_max
                    # Ruptura Tekla: 1.0 mm (2.835 pt) în șabloane suedeze sau 3.0 mm (8.5 pt) în cele norvegiene, până la 16 pt
                    if 1.2 <= gap <= 16.0:
                        # Verificăm ca niciun capăt al rupturii să nu aibă alte linii sau curbe conectate.
                        # Într-o ruptură Tekla reală, capetele de linie sunt libere în aer.
                        p_left = fitz.Point(x1_max, y1)
                        p_right = fitz.Point(x2_min, y2)
                        has_conn = False
                        for other_idx, it in enumerate(items):
                            if other_idx in (idx1, idx2):
                                continue
                            if it[0] == "l":
                                if math.hypot(it[1].x - p_left.x, it[1].y - p_left.y) < 0.5 or math.hypot(it[2].x - p_left.x, it[2].y - p_left.y) < 0.5:
                                    has_conn = True
                                    break
                                if math.hypot(it[1].x - p_right.x, it[1].y - p_right.y) < 0.5 or math.hypot(it[2].x - p_right.x, it[2].y - p_right.y) < 0.5:
                                    has_conn = True
                                    break
                            elif it[0] == "c":
                                if math.hypot(it[1].x - p_left.x, it[1].y - p_left.y) < 0.5 or math.hypot(it[4].x - p_left.x, it[4].y - p_left.y) < 0.5:
                                    has_conn = True
                                    break
                                if math.hypot(it[1].x - p_right.x, it[1].y - p_right.y) < 0.5 or math.hypot(it[4].x - p_right.x, it[4].y - p_right.y) < 0.5:
                                    has_conn = True
                                    break
                        if not has_conn:
                            break_pairs.append((x1_max, x2_min, y1, idx1, idx2))

        if not break_pairs:
            return []

        # Grupăm perechile de ruptură după coordonata centrală X
        clusters = []
        for xl, xr, y, idx1, idx2 in sorted(break_pairs, key=lambda x: x[0]):
            center = (xl + xr) / 2.0
            matched = False
            for c in clusters:
                if abs(c["center"] - center) < 3.0:
                    c["pairs"].append((xl, xr, y, idx1, idx2))
                    c["xl"] = min(c["xl"], xl)
                    c["xr"] = max(c["xr"], xr)
                    matched = True
                    break
            if not matched:
                clusters.append({
                    "center": center,
                    "xl": xl,
                    "xr": xr,
                    "pairs": [(xl, xr, y, idx1, idx2)]
                })

        # O ruptură Tekla validă taie secțiunea pe verticală (cel puțin 2 perechi aliniate: latura de sus și latura de jos)
        clusters = [c for c in clusters if len(c["pairs"]) >= 2]
        clusters.sort(key=lambda c: c["center"])
        return clusters

    @classmethod
    def detect_holes_in_bbox(cls, page, bbox):
        """
        Detectează găurile circulare ale piesei aflate în interiorul vederii Front.
        În PDF-urile Tekla, găurile sunt reprezentate prin 4 curbe Bezier circulare.
        Returnează lista [{'center': (cx, cy), 'radius': r_pt}]
        """
        drawings = page.get_drawings()
        holes = []
        for d in drawings:
            r = d["rect"]
            # Verificăm să fie în interiorul cutiei de încadrare a vederii Front
            if not (bbox["min_x"] - 5.0 <= r.x0 and r.x1 <= bbox["max_x"] + 5.0 and
                    bbox["min_y"] - 5.0 <= r.y0 and r.y1 <= bbox["max_y"] + 5.0):
                continue

            items = d.get("items", [])
            n_curves = sum(1 for it in items if it[0] == "c")
            if (n_curves == 4 and len(items) <= 6) or (len(items) == 6 and n_curves == 6):
                cx = (r.x0 + r.x1) / 2.0
                cy = (r.y0 + r.y1) / 2.0
                rx = (r.x1 - r.x0) / 2.0
                ry = (r.y1 - r.y0) / 2.0
                min_r = 0.8 if len(items) > 4 else 0.4
                if rx > 0 and ry > 0 and abs(rx - ry) < 0.35 * max(rx, ry) and min_r <= rx <= 45.0:
                    rad = min(rx, ry) if n_curves == 6 else (rx + ry) / 2.0
                    if not any(math.hypot(cx - h["center"][0], cy - h["center"][1]) < 1.0 for h in holes):
                        holes.append({
                            "center": (cx, cy),
                            "radius": rad
                        })
        return holes

    @classmethod
    def _calculate_break_expansions(cls, stations, items, bbox, target_len, base_scale, page, stykk_y, holes=None):
        """
        Calculează expansiunea exactă în milimetri pentru fiecare stație de ruptură (break station).
        Pentru 1 ruptură: exp = target_len - total_drawn_solid_mm.
        Pentru multi-break (N rupturi): asociază direct liniile orizontale de cotă din PDF cu
        stațiile de ruptură pe care le acoperă și rezolvă sistemul exact fără a distorsiona cotele găurilor.
        """
        total_gaps_pt = sum(s["xr"] - s["xl"] for s in stations)
        total_drawn_solid_mm = (bbox["width"] - total_gaps_pt) * base_scale
        total_missing_mm = target_len - total_drawn_solid_mm

        if len(stations) == 1:
            stations[0]["exp"] = max(0.0, total_missing_mm)
            return stations

        words = page.get_text("words")
        max_dim_y = (stykk_y - 5.0) if stykk_y else (page.rect.height * 0.70)

        # 1. Metoda exactă: extragere linii orizontale de cotă [xl, xr] și rezolvare sistem pe stații
        h_dim_lines = []
        for d in page.get_drawings():
            w = d.get("width")
            if not w or w > 0.52:
                continue
            for it in d["items"]:
                if it[0] == "l":
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 0.25 and abs(p1.x - p2.x) > 2.0:
                        xl, xr = min(p1.x, p2.x), max(p1.x, p2.x)
                        yl = (p1.y + p2.y) / 2.0
                        if bbox["min_x"] - 8.0 <= xl and xr <= bbox["max_x"] + 8.0:
                            if yl < bbox["min_y"] - 2.0 or yl > bbox["max_y"] + 2.0:
                                h_dim_lines.append((xl, xr, yl))

        equations = []
        dim_words = []
        for w in words:
            txt = w[4].strip().replace(",", ".")
            if txt.replace(".", "", 1).isdigit() and w[1] < max_dim_y:
                val = float(txt)
                if val <= 10:
                    continue
                xm = (w[0] + w[2]) / 2.0
                y1 = w[3]
                if val < target_len - 10.0:
                    dim_words.append((w[0], w[2], xm, w[1], val))
                cands = [
                    (xl, xr, yl) for (xl, xr, yl) in h_dim_lines
                    if (y1 - 1.5 <= yl <= y1 + 2.5) and (xl - 5.0 <= xm <= xr + 5.0)
                ]
                if cands:
                    xl, xr, yl = min(cands, key=lambda c: abs(c[2] - y1) + abs((c[0] + c[1]) / 2.0 - xm) * 0.1)
                    covered = tuple(
                        j for j, st in enumerate(stations)
                        if xl - 0.8 <= st["xl"] and st["xr"] <= xr + 0.8
                    )
                    if covered:
                        gaps_s = sum(stations[j]["xr"] - stations[j]["xl"] for j in covered)
                        seg_solid_mm = ((xr - xl) - gaps_s) * base_scale
                        rhs = max(0.0, val - seg_solid_mm)
                        equations.append((len(covered), abs((xl + xr) / 2.0 - xm), covered, rhs))

        if equations:
            equations.append((len(stations), 9999.0, tuple(range(len(stations))), max(0.0, total_missing_mm)))
            equations.sort(key=lambda e: (e[0], e[1]))
            exps = [None] * len(stations)
            changed = True
            while changed:
                changed = False
                for _, _, covered, rhs in equations:
                    unknowns = [j for j in covered if exps[j] is None]
                    if len(unknowns) == 1:
                        u = unknowns[0]
                        known_sum = sum(exps[j] for j in covered if j != u)
                        if rhs >= known_sum - 1.0:
                            exps[u] = max(0.0, rhs - known_sum)
                            changed = True
                            break

            if all(e is not None for e in exps) and abs(total_missing_mm - sum(exps)) <= max(10.0, 2.0 * len(stations)):
                for i, s in enumerate(stations):
                    s["exp"] = exps[i]
                return stations

        # 2. Fallback: linii verticale, capete de decupaje pe margini și găuri pentru delimitarea segmentelor
        feat_x = []
        for it in items:
            if it[0] == "l" and abs(it[1].x - it[2].x) < 0.25:
                vx = round((it[1].x + it[2].x) / 2.0, 2)
                if bbox["min_x"] + 1.0 < vx < bbox["max_x"] - 1.0:
                    feat_x.append(vx)
            elif it[0] == "l" and abs(it[1].y - it[2].y) < 0.25:
                ym = (it[1].y + it[2].y) / 2.0
                if abs(ym - bbox["min_y"]) < 0.6 or abs(ym - bbox["max_y"]) < 0.6:
                    for x_end in (it[1].x, it[2].x):
                        vx = round(x_end, 2)
                        if bbox["min_x"] + 1.0 < vx < bbox["max_x"] - 1.0:
                            if not any(abs(vx - st["xl"]) < 1.0 or abs(vx - st["xr"]) < 1.0 for st in stations):
                                feat_x.append(vx)

        if holes:
            for h in holes:
                hx = round(h["center"][0], 2)
                if bbox["min_x"] + 1.0 < hx < bbox["max_x"] - 1.0:
                    feat_x.append(hx)
        feat_x = sorted(list(set(feat_x)))

        # Stabilim granițele fizice [b_left, b_right] pentru fiecare stație de ruptură în parte.
        # Decupajele sau caracteristicile dintre două stații nu aparțin niciunei stații și rămân nedistorsionate.
        station_bounds = []
        for i, s in enumerate(stations):
            if i == 0:
                left_feats = [fx for fx in feat_x if fx <= s["xl"]]
                b_left = max(left_feats) if left_feats else bbox["min_x"]
            else:
                between_prev = [fx for fx in feat_x if stations[i - 1]["xr"] <= fx <= s["xl"]]
                if between_prev:
                    b_left = max(between_prev)
                else:
                    b_left = (stations[i - 1]["xr"] + s["xl"]) / 2.0

            if i == len(stations) - 1:
                right_feats = [fx for fx in feat_x if fx >= s["xr"]]
                b_right = min(right_feats) if right_feats else bbox["max_x"]
            else:
                between_next = [fx for fx in feat_x if s["xr"] <= fx <= stations[i + 1]["xl"]]
                if between_next:
                    b_right = min(between_next)
                else:
                    b_right = (s["xr"] + stations[i + 1]["xl"]) / 2.0

            station_bounds.append((b_left, b_right))

        matched_dims = []
        for i, s in enumerate(stations):
            b_left, b_right = station_bounds[i]
            candidates = [
                dw for dw in dim_words
                if (b_left - 5.0 <= dw[2] <= b_right + 5.0) or abs(dw[2] - s["center"]) < 15.0
            ]
            if candidates:
                best = min(candidates, key=lambda x: abs(x[2] - s["center"]))
                matched_dims.append(best[4])
            else:
                matched_dims.append(None)

        expansions = []
        if all(d is not None for d in matched_dims):
            for i in range(len(stations)):
                b_left, b_right = station_bounds[i]
                gap = stations[i]["xr"] - stations[i]["xl"]
                seg_drawn_mm = ((b_right - b_left) - gap) * base_scale
                exp_i = max(0.0, matched_dims[i] - seg_drawn_mm)
                expansions.append(exp_i)

            sum_exp = sum(expansions)
            if sum_exp > 0:
                scale_exp = total_missing_mm / sum_exp
                for i, s in enumerate(stations):
                    s["exp"] = expansions[i] * scale_exp
                return stations

        # Fallback uniform dacă nu toate cotele au fost găsite sau sum_exp <= 0
        exp_per_station = total_missing_mm / len(stations)
        for s in stations:
            s["exp"] = max(0.0, exp_per_station)

        return stations

    @staticmethod
    def _merge_collinear_lines(cad_lines):
        """
        Unește segmentele coliniare adiacente sau suprapuse într-o singură linie continuă.
        Elimină micro-segmentele, golurile și liniile duble pentru margini 100% curate în CAD.
        """
        if not cad_lines:
            return []

        h_by_y = defaultdict(list)
        v_by_x = defaultdict(list)
        others = []

        for p1, p2 in cad_lines:
            x1, y1 = p1[0], p1[1]
            x2, y2 = p2[0], p2[1]
            if abs(y1 - y2) < 0.25:
                y_key = round((y1 + y2) / 2.0, 2)
                h_by_y[y_key].append((min(x1, x2), max(x1, x2)))
            elif abs(x1 - x2) < 0.25:
                x_key = round((x1 + x2) / 2.0, 2)
                v_by_x[x_key].append((min(y1, y2), max(y1, y2)))
            else:
                others.append((p1, p2))

        merged = []

        # Unire orizontale
        for y, intervals in h_by_y.items():
            intervals.sort(key=lambda iv: iv[0])
            cur_x0, cur_x1 = intervals[0]
            for x0, x1 in intervals[1:]:
                if x0 <= cur_x1 + 0.5:  # contiguous or overlapping
                    cur_x1 = max(cur_x1, x1)
                else:
                    if cur_x1 - cur_x0 > 0.05:
                        merged.append(((cur_x0, y), (cur_x1, y)))
                    cur_x0, cur_x1 = x0, x1
            if cur_x1 - cur_x0 > 0.05:
                merged.append(((cur_x0, y), (cur_x1, y)))

        # Unire verticale
        for x, intervals in v_by_x.items():
            intervals.sort(key=lambda iv: iv[0])
            cur_y0, cur_y1 = intervals[0]
            for y0, y1 in intervals[1:]:
                if y0 <= cur_y1 + 0.5:
                    cur_y1 = max(cur_y1, y1)
                else:
                    if cur_y1 - cur_y0 > 0.05:
                        merged.append(((x, cur_y0), (x, cur_y1)))
                    cur_y0, cur_y1 = y0, y1
            if cur_y1 - cur_y0 > 0.05:
                merged.append(((x, cur_y0), (x, cur_y1)))

        merged.extend(others)
        return merged

    @staticmethod
    def _fit_circle_3pts(p1, p2, p3):
        """Calculează centrul și raza cercului care trece prin 3 puncte necoliniare."""
        x1, y1 = p1
        x2, y2 = p2
        x3, y3 = p3
        D = 2 * (x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2))
        if abs(D) < 1e-7:
            return None
        cx = ((x1**2 + y1**2) * (y2 - y3) + (x2**2 + y2**2) * (y3 - y1) + (x3**2 + y3**2) * (y1 - y2)) / D
        cy = ((x1**2 + y1**2) * (x3 - x2) + (x2**2 + y2**2) * (x1 - x3) + (x3**2 + y3**2) * (x2 - x1)) / D
        R = math.sqrt((x1 - cx)**2 + (y1 - cy)**2)
        return cx, cy, R

    @staticmethod
    def _get_arc_angles(cx, cy, pa, pmid, pb):
        """Găsește unghiul de start și de sfârșit pentru un arc DXF trasat CCW."""
        a = math.degrees(math.atan2(pa[1] - cy, pa[0] - cx)) % 360
        b = math.degrees(math.atan2(pb[1] - cy, pb[0] - cx)) % 360
        m = math.degrees(math.atan2(pmid[1] - cy, pmid[0] - cx)) % 360
        sweep_ab = (b - a) % 360
        sweep_am = (m - a) % 360
        if 0 < sweep_am < sweep_ab:
            return a, b
        else:
            return b, a

    @staticmethod
    def calculate_arc_bulge(p1, p2, center, mid):
        """
        Calculează valoarea bulge cu semnul matematic corect pentru un arc de la p1 la p2
        care trece prin mid, având centrul în center.
        bulge = tan(theta / 4)
        bulge > 0: arc CCW (la stânga)
        bulge < 0: arc CW (la dreapta)
        """
        cx, cy = center
        a1 = math.degrees(math.atan2(p1[1] - cy, p1[0] - cx)) % 360
        a2 = math.degrees(math.atan2(p2[1] - cy, p2[0] - cx)) % 360
        am = math.degrees(math.atan2(mid[1] - cy, mid[0] - cx)) % 360

        sweep_ccw = (a2 - a1) % 360
        mid_ccw = (am - a1) % 360

        if 0 < mid_ccw < sweep_ccw:
            theta_rad = math.radians(sweep_ccw)
            return math.tan(theta_rad / 4.0)
        else:
            sweep_cw = (a1 - a2) % 360
            theta_rad = math.radians(sweep_cw)
            return -math.tan(theta_rad / 4.0)

    @staticmethod
    def snap_points_grid(segments, tol=0.5):
        """
        Snap-uiește punctele folosind o grilă de discretizare spațială (spatial hash) și Union-Find.
        Garantează fuziunea tranzitivă completă a nodurilor adiacente din toleranță.
        Elimină 100% micro-golurile (gaps) și treptele dintre linii.
        """
        if not segments:
            return []

        raw_pts = []
        for p1, p2 in segments:
            raw_pts.append((float(p1[0]), float(p1[1])))
            raw_pts.append((float(p2[0]), float(p2[1])))

        parent = list(range(len(raw_pts)))

        def find(i):
            path = []
            while parent[i] != i:
                path.append(i)
                i = parent[i]
            for node in path:
                parent[node] = i
            return i

        def union(i, j):
            root_i = find(i)
            root_j = find(j)
            if root_i != root_j:
                parent[root_j] = root_i

        grid = defaultdict(list)
        cell_size = tol
        for idx, (x, y) in enumerate(raw_pts):
            cx = int(math.floor(x / cell_size))
            cy = int(math.floor(y / cell_size))
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    cell_key = (cx + dx, cy + dy)
                    if cell_key in grid:
                        for other_idx in grid[cell_key]:
                            ox, oy = raw_pts[other_idx]
                            if math.hypot(x - ox, y - oy) <= tol:
                                union(idx, other_idx)
            grid[(cx, cy)].append(idx)

        groups = defaultdict(list)
        for idx in range(len(raw_pts)):
            groups[find(idx)].append(raw_pts[idx])

        group_centers = {}
        for root, pts_in_group in groups.items():
            avg_x = sum(p[0] for p in pts_in_group) / float(len(pts_in_group))
            avg_y = sum(p[1] for p in pts_in_group) / float(len(pts_in_group))
            rx = round(avg_x, 2)
            ry = round(avg_y, 2)
            if abs(rx - round(rx)) < 0.03:
                rx = round(rx)
            if abs(ry - round(ry)) < 0.03:
                ry = round(ry)
            group_centers[root] = (rx, ry)

        snapped_segments = []
        for i in range(0, len(raw_pts), 2):
            sp1 = group_centers[find(i)]
            sp2 = group_centers[find(i + 1)]
            if sp1 != sp2:
                snapped_segments.append((sp1, sp2))

        return snapped_segments

    @classmethod
    def extract_all_closed_loops(cls, lines, arcs=None):
        """
        Extrage toate ciclurile închise (contur exterior + decupaje / găuri interioare).
        Fiecare ciclu este o listă ordonată de vertexuri (x, y, 0, 0, bulge) gata pentru LWPOLYLINE.
        Elimină micro-liniile parazite agățate de colțuri și închide automat eventualele
        goluri minore rămase la exportul PDF (< 25 mm) pentru a garanta contururi 100% închise.
        """
        adj = defaultdict(list)
        edges = []

        for p1, p2 in lines:
            p1_r = (round(p1[0], 2), round(p1[1], 2))
            p2_r = (round(p2[0], 2), round(p2[1], 2))
            if p1_r == p2_r:
                continue
            e_idx = len(edges)
            edges.append((p1_r, p2_r, False))
            adj[p1_r].append((p2_r, 0.0, e_idx))
            adj[p2_r].append((p1_r, 0.0, e_idx))

        if arcs:
            for p_start, p_end, center, mid in arcs:
                s_r = (round(p_start[0], 2), round(p_start[1], 2))
                e_r = (round(p_end[0], 2), round(p_end[1], 2))
                if s_r == e_r:
                    continue
                b_fwd = cls.calculate_arc_bulge(s_r, e_r, center, mid)
                b_rev = cls.calculate_arc_bulge(e_r, s_r, center, mid)
                e_idx = len(edges)
                edges.append((s_r, e_r, True))
                adj[s_r].append((e_r, b_fwd, e_idx))
                adj[e_r].append((s_r, b_rev, e_idx))

        visited_edges = set()

        # Colectăm componentele conexe ale grafului
        all_comps = []
        visited_comp_nodes = set()
        for start_node in list(adj.keys()):
            if start_node in visited_comp_nodes:
                continue
            comp_nodes = set()
            comp_edges = set()
            stack = [start_node]
            visited_comp_nodes.add(start_node)
            while stack:
                curr = stack.pop()
                comp_nodes.add(curr)
                for nxt, _, e_idx in adj[curr]:
                    if e_idx in visited_edges:
                        continue
                    comp_edges.add(e_idx)
                    if nxt not in visited_comp_nodes:
                        visited_comp_nodes.add(nxt)
                        stack.append(nxt)
            all_comps.append((comp_nodes, comp_edges))

        # Eliminăm componentele parazite exterioare minuscule (ex. săgeata de sudură din HSQP.1018)
        if len(all_comps) > 1:
            comp_bboxes = []
            for c_nodes, c_edges in all_comps:
                cxs = [n[0] for n in c_nodes]
                cys = [n[1] for n in c_nodes]
                cx0, cx1 = min(cxs), max(cxs)
                cy0, cy1 = min(cys), max(cys)
                comp_bboxes.append((cx0, cy0, cx1, cy1, (cx1 - cx0) + (cy1 - cy0)))
            main_idx = max(range(len(comp_bboxes)), key=lambda i: comp_bboxes[i][4])
            mx0, my0, mx1, my1, m_span = comp_bboxes[main_idx]
            for i, (c_nodes, c_edges) in enumerate(all_comps):
                if i == main_idx:
                    continue
                cx0, cy0, cx1, cy1, c_span = comp_bboxes[i]
                outside_main = (
                    cx1 < mx0 - 2.0 or cx0 > mx1 + 2.0 or
                    cy1 < my0 - 2.0 or cy0 > my1 + 2.0
                )
                if outside_main and c_span < 0.25 * m_span:
                    visited_edges.update(c_edges)

        # Eliminăm ramurile deschise (mustăți/teșituri oprite în T-joncțiuni):
        # - Pe componentele conexe care conțin cicluri (n_edges >= n_nodes),
        #   reducem la 2-core eliminând iterativ orice nod de grad 1 (inclusiv lanțuri prin noduri de grad 2).
        # - Pe componentele fără cicluri (arbori, ex. contur cu mic gap de închidere),
        #   eliminăm doar mustățile de grad 1 lipite direct de un nod cu grad >= 3.
        for comp_nodes, comp_edges in all_comps:
            if not comp_edges or comp_edges.issubset(visited_edges):
                continue
            has_cycle = len(comp_edges) >= len(comp_nodes)
            changed = True
            while changed:
                changed = False
                for node in list(comp_nodes):
                    active = [c for c in adj[node] if c[2] not in visited_edges]
                    if len(active) == 1:
                        nxt, _, e_idx = active[0]
                        if has_cycle:
                            visited_edges.add(e_idx)
                            changed = True
                        else:
                            nxt_active = [c for c in adj[nxt] if c[2] not in visited_edges]
                            if len(nxt_active) >= 3:
                                visited_edges.add(e_idx)
                                changed = True

        loops = []

        # Faza 1: Procesăm mai întâi lanțurile care au capete deschise (deg == 1),
        # astfel încât să fie parcurse cap-coadă și închise dacă golul <= 25 mm
        deg1_nodes = [
            n for n in sorted(adj.keys())
            if len([c for c in adj[n] if c[2] not in visited_edges]) == 1
        ]
        for node in deg1_nodes:
            active_here = [c for c in adj[node] if c[2] not in visited_edges]
            if not active_here:
                continue

            chain = []
            curr = node
            start_loop = node
            while True:
                candidates = [c for c in adj[curr] if c[2] not in visited_edges]
                if not candidates:
                    break
                nxt, b_val, edge_num = candidates[0]
                visited_edges.add(edge_num)
                chain.append((curr[0], curr[1], 0, 0, b_val))
                curr = nxt
                if curr == start_loop:
                    break

            if curr == start_loop and len(chain) >= 3:
                loops.append(chain)
            elif len(chain) >= 2:
                gap_dist = math.hypot(curr[0] - start_loop[0], curr[1] - start_loop[1])
                if gap_dist <= 25.0:
                    chain.append((curr[0], curr[1], 0, 0, 0.0))
                    if len(chain) >= 3:
                        loops.append(chain)

        # Faza 2: Extragem toate ciclurile complet închise (deg >= 2) prin căutare DFS de ciclu simplu
        def _find_simple_cycle(start_node, first_nxt, first_bulge, first_e_idx):
            steps = [0]

            def _dfs(curr_node, chain_acc, edge_acc, vis_nodes):
                local_added_nodes = []
                local_chain = list(chain_acc)
                local_edges = list(edge_acc)
                curr = curr_node

                while True:
                    steps[0] += 1
                    if steps[0] > 50000:
                        for n in local_added_nodes:
                            vis_nodes.remove(n)
                        return None

                    cands = [
                        c for c in adj[curr]
                        if c[2] not in visited_edges and c[2] != first_e_idx
                        and (c[0] == start_node or c[0] not in vis_nodes)
                    ]
                    if not cands:
                        for n in local_added_nodes:
                            vis_nodes.remove(n)
                        return None

                    if len(local_chain) >= 2:
                        cands.sort(key=lambda c: 0 if c[0] == start_node else 1)

                    if len(cands) == 1:
                        nxt, b_val, edge_num = cands[0]
                        if nxt == start_node:
                            if len(local_chain) >= 2:
                                return (
                                    local_chain + [(curr[0], curr[1], 0, 0, b_val)],
                                    local_edges + [edge_num],
                                )
                            for n in local_added_nodes:
                                vis_nodes.remove(n)
                            return None
                        local_chain.append((curr[0], curr[1], 0, 0, b_val))
                        local_edges.append(edge_num)
                        vis_nodes.add(nxt)
                        local_added_nodes.append(nxt)
                        curr = nxt
                    else:
                        for nxt, b_val, edge_num in cands:
                            if nxt == start_node:
                                if len(local_chain) >= 2:
                                    return (
                                        local_chain + [(curr[0], curr[1], 0, 0, b_val)],
                                        local_edges + [edge_num],
                                    )
                                continue
                            vis_nodes.add(nxt)
                            res = _dfs(
                                nxt,
                                local_chain + [(curr[0], curr[1], 0, 0, b_val)],
                                local_edges + [edge_num],
                                vis_nodes,
                            )
                            vis_nodes.remove(nxt)
                            if res is not None:
                                return res
                        for n in local_added_nodes:
                            vis_nodes.remove(n)
                        return None

            return _dfs(
                first_nxt,
                [(start_node[0], start_node[1], 0, 0, first_bulge)],
                [first_e_idx],
                {first_nxt},
            )

        for node in sorted(adj.keys()):
            for next_pt, bulge, e_idx in adj[node]:
                if e_idx in visited_edges:
                    continue

                res = _find_simple_cycle(node, next_pt, bulge, e_idx)
                if res is not None:
                    chain, c_edges = res
                    if len(chain) >= 3:
                        loops.append(chain)
                        visited_edges.update(c_edges)
                        c_nodes = {(pt[0], pt[1]) for pt in chain}
                        for v_node in c_nodes:
                            for nxt, _, cord_eid in adj[v_node]:
                                if cord_eid not in visited_edges and nxt in c_nodes:
                                    visited_edges.add(cord_eid)

        return loops

    @classmethod
    def _detect_fillet_arcs(cls, items, exclude_indices=None):
        """
        Identifică arcele de cerc (fillet) reprezentate în PDF prin secvențe
        de micro-linii tangente și curbe Bezier cubice (l + c + l sau doar c).
        Găsește centrul real, raza exactă și unghiurile arcului.
        """
        if exclude_indices is None:
            exclude_indices = set()

        curves = [(idx, it) for idx, it in enumerate(items) if it[0] == "c"]
        lines = [(idx, it) for idx, it in enumerate(items) if it[0] == "l"]

        fillet_used_indices = set()
        fillet_arcs = []

        for c_idx, c_it in curves:
            if c_idx in exclude_indices:
                continue
            p1, p2, p3, p4 = c_it[1], c_it[2], c_it[3], c_it[4]

            adj_p1 = None
            adj_p4 = None
            for l_idx, l_it in lines:
                if l_idx in exclude_indices:
                    continue
                lp1, lp2 = l_it[1], l_it[2]
                l_len = math.hypot(lp1.x - lp2.x, lp1.y - lp2.y)
                if l_len <= 1.2:
                    if math.hypot(lp1.x - p1.x, lp1.y - p1.y) < 0.08:
                        adj_p1 = (l_idx, lp2)
                    elif math.hypot(lp2.x - p1.x, lp2.y - p1.y) < 0.08:
                        adj_p1 = (l_idx, lp1)
                    elif math.hypot(lp1.x - p4.x, lp1.y - p4.y) < 0.08:
                        adj_p4 = (l_idx, lp2)
                    elif math.hypot(lp2.x - p4.x, lp2.y - p4.y) < 0.08:
                        adj_p4 = (l_idx, lp1)

            start_pt = adj_p1[1] if adj_p1 else fitz.Point(p1.x, p1.y)
            end_pt = adj_p4[1] if adj_p4 else fitz.Point(p4.x, p4.y)

            t = 0.5
            bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
            by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
            mid_pt = fitz.Point(bx, by)

            fit = cls._fit_circle_3pts((start_pt.x, start_pt.y), (mid_pt.x, mid_pt.y), (end_pt.x, end_pt.y))
            if fit:
                cx_pdf, cy_pdf, r_pdf = fit
                d_p1 = math.hypot(p1.x - cx_pdf, p1.y - cy_pdf)
                d_p4 = math.hypot(p4.x - cx_pdf, p4.y - cy_pdf)
                if abs(d_p1 - r_pdf) / r_pdf < 0.05 and abs(d_p4 - r_pdf) / r_pdf < 0.05:
                    fillet_arcs.append((cx_pdf, cy_pdf, r_pdf, start_pt, mid_pt, end_pt))
                    fillet_used_indices.add(c_idx)
                    if adj_p1:
                        fillet_used_indices.add(adj_p1[0])
                    if adj_p4:
                        fillet_used_indices.add(adj_p4[0])

        return fillet_arcs, fillet_used_indices

    def convert_single_pdf(
        self,
        pdf_path,
        output_dir,
        output_format="dwg",
        only_front_view=True,
        calibrate_1to1=True,
        custom_name=None,
    ):
        """
        Convertește un singur PDF într-un fișier CAD (.dwg sau .dxf) 1:1.
        Păstrează decupajele și găurile 1:1 nedistorsionate și extinde vederile scurtate la lungimea reală.
        """
        if not self.is_available():
            return False, None, "Bibliotecile PyMuPDF sau ezdxf nu sunt instalate."

        pdf_path = str(pdf_path)
        if not os.path.isfile(pdf_path):
            return False, None, f"Fișierul PDF nu există: {pdf_path}"

        try:
            doc = fitz.open(pdf_path)
            if len(doc) == 0:
                doc.close()
                return False, None, "Fișierul PDF nu conține pagini."
            page = doc[0]

            # 1. Extragere metadate desen
            meta = self.extract_drawing_metadata(page, pdf_path)
            clean_name = custom_name if custom_name else meta["clean_name"]

            # 2. Extragere geometrie curată
            items, bbox = self.extract_geometry_items(
                page, stykk_y=meta["stykk_y"], top_margin_y=meta.get("top_margin_y", 20.0), only_front_view=only_front_view, meta=meta
            )

            if not items or not bbox or bbox["width"] <= 0 or bbox["height"] <= 0:
                doc.close()
                return False, None, "Nu s-a putut izola conturul piesei (lipsă vectori sau desen raster)."

            base_scale_factor = (25.4 / 72.0) * meta["scale"]
            min_x = bbox["min_x"]
            max_y = bbox["max_y"]

            drawn_w_pt = bbox["width"]
            drawn_h_pt = bbox["height"]
            drawn_len_mm = drawn_w_pt * base_scale_factor

            target_len = meta["lengde"]
            target_wid = meta["profil_width"]

            # 3. Detectare stații de ruptură Tekla (view shortening breaks)
            stations = []
            is_shortened = False
            raw_stations = self._find_view_break_stations(items, bbox)

            # Dacă avem ruptură Tekla și lățimea reală a piesei (target_wid), calibrăm întâi scara de bază pe Y
            # (necesar pentru desenele cu Scale: INGEN, ex. 1:14.142 în loc de 1:10)
            if calibrate_1to1 and raw_stations and target_wid:
                h_mm_at_base = drawn_h_pt * base_scale_factor
                if h_mm_at_base > 1.0:
                    wid_ratio = target_wid / h_mm_at_base
                    if 0.25 <= wid_ratio <= 2.5 and abs(wid_ratio - 1.0) > 0.03:
                        base_scale_factor *= wid_ratio
                        drawn_len_mm = drawn_w_pt * base_scale_factor

            total_break_gap_pt = sum((s["xr"] - s["xl"]) for s in raw_stations)
            drawn_metal_mm = max(0.0, (drawn_w_pt - total_break_gap_pt) * base_scale_factor)

            # Dacă avem ruptură Tekla și lungimea extrasă este mai mică decât metalul desenat efectiv, a fost citit un ANTALL sau alt număr
            if raw_stations and target_len and target_len <= drawn_metal_mm * 0.98:
                target_len = None

            # Dacă avem ruptură Tekla dar nu avem target_len valid, căutăm cota principală lângă vederea Front
            if calibrate_1to1 and raw_stations and not target_len:
                dim_candidates = []
                for w in page.get_text("words"):
                    clean_w = w[4].replace(",", ".")
                    if clean_w.replace(".", "", 1).isdigit():
                        v = float(clean_w)
                        if drawn_metal_mm * 1.02 < v < 50000:
                            dim_candidates.append((w[1], v))
                if dim_candidates:
                    dim_candidates.sort(key=lambda item: abs(item[0] - bbox["min_y"]))
                    target_len = dim_candidates[0][1]
                    meta["lengde"] = target_len

            # Detectare găuri circulare în interiorul vederii Front
            holes = self.detect_holes_in_bbox(page, bbox)

            if calibrate_1to1 and target_len and raw_stations and (target_len > drawn_metal_mm + 1.0):
                stations = self._calculate_break_expansions(
                    raw_stations, items, bbox, target_len, base_scale_factor, page, meta["stykk_y"], holes=holes
                )
                is_shortened = bool(stations)

            doc.close()

            # 4. Funcție de transformare coordonate din PDF în CAD (milimetri 1:1)
            if is_shortened:
                def pt_to_cad(pt):
                    x = pt.x
                    cad_x = (x - min_x) * base_scale_factor
                    for s in stations:
                        bl, br, exp = s["xl"], s["xr"], s["exp"]
                        if x > bl:
                            gap_pts = min(x, br) - bl
                            cad_x -= gap_pts * base_scale_factor
                            cad_x += exp
                    cad_y = (max_y - pt.y) * base_scale_factor
                    return (cad_x, cad_y)

                dim_x_final = target_len
                dim_y_final = drawn_h_pt * base_scale_factor

            else:
                scale_final = base_scale_factor
                if calibrate_1to1 and target_len and target_wid:
                    target_max = max(target_len, target_wid)
                    target_min = min(target_len, target_wid)
                    curr_max_mm = max(drawn_w_pt, drawn_h_pt) * base_scale_factor
                    curr_min_mm = min(drawn_w_pt, drawn_h_pt) * base_scale_factor
                    if curr_max_mm > 0 and curr_min_mm > 0 and target_min > 0:
                        ratio = target_max / curr_max_mm
                        drawn_aspect = curr_max_mm / curr_min_mm
                        target_aspect = target_max / target_min
                        if 0.85 <= ratio <= 1.15:
                            scale_final = base_scale_factor * ratio
                        elif (0.85 <= (drawn_aspect / target_aspect) <= 1.15) and (0.20 <= ratio <= 4.0):
                            # Cazul desenelor fără scară numerică (ex. Scale: INGEN) cu raport de aspect corect
                            scale_final = base_scale_factor * ratio

                def pt_to_cad(pt):
                    return (
                        (pt.x - min_x) * scale_final,
                        (max_y - pt.y) * scale_final
                    )

                dim_x_final = drawn_w_pt * scale_final
                dim_y_final = drawn_h_pt * scale_final

            # 5. Identificare arce de racordare (fillet) pentru export nativ ARC
            fillet_arcs, fillet_used_indices = self._detect_fillet_arcs(items)

            # 6. Colectare linii CAD brute (din PDF + linii de unire peste rupturile Tekla)
            raw_cad_lines = []
            for idx, it in enumerate(items):
                if idx in fillet_used_indices:
                    continue

                cmd = it[0]
                if cmd == "l":
                    p1 = pt_to_cad(it[1])
                    p2 = pt_to_cad(it[2])
                    raw_cad_lines.append((p1, p2))

                elif cmd == "c":
                    p1, p2, p3, p4 = it[1], it[2], it[3], it[4]
                    steps = 16
                    pts = []
                    for s in range(steps + 1):
                        t = s / float(steps)
                        bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
                        by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
                        pts.append(pt_to_cad(fitz.Point(bx, by)))
                    for i in range(len(pts) - 1):
                        raw_cad_lines.append((pts[i], pts[i + 1]))

                elif cmd == "re":
                    r = it[1]
                    p0 = pt_to_cad(fitz.Point(r.x0, r.y0))
                    p1 = pt_to_cad(fitz.Point(r.x1, r.y0))
                    p2 = pt_to_cad(fitz.Point(r.x1, r.y1))
                    p3 = pt_to_cad(fitz.Point(r.x0, r.y1))
                    raw_cad_lines.extend([(p0, p1), (p1, p2), (p2, p3), (p3, p0)])

            # Adăugăm liniile de legătură (bridging lines) peste fiecare stație de ruptură
            if is_shortened:
                for s in stations:
                    for xl, xr, y, idx1, idx2 in s["pairs"]:
                        p_left = pt_to_cad(fitz.Point(xl, y))
                        p_right = pt_to_cad(fitz.Point(xr, y))
                        raw_cad_lines.append((p_left, p_right))
            elif raw_stations:
                # Chiar dacă nu se extinde scara, unim rupturile Tekla pentru a nu lăsa goluri deschise în contur
                for s in raw_stations:
                    for xl, xr, y, idx1, idx2 in s["pairs"]:
                        p_left = pt_to_cad(fitz.Point(xl, y))
                        p_right = pt_to_cad(fitz.Point(xr, y))
                        raw_cad_lines.append((p_left, p_right))

            # 7. Unire muchii coliniare adiacente (eliminare linii suprapuse)
            merged_lines = self._merge_collinear_lines(raw_cad_lines)

            # 8. Snapping magnetic de noduri (eliminare micro-goluri 0.01 - 0.5 mm)
            final_lines = self.snap_points_grid(merged_lines, tol=0.5)

            # 9. Snap capete arce la cele mai apropiate noduri din contur
            snapped_arcs_data = []
            if fillet_arcs:
                pts_contour = [p for l in final_lines for p in l]
                for cx_pdf, cy_pdf, r_pdf, start_pt, mid_pt, end_pt in fillet_arcs:
                    c_cad = pt_to_cad(fitz.Point(cx_pdf, cy_pdf))
                    start_cad = pt_to_cad(start_pt)
                    end_cad = pt_to_cad(end_pt)
                    mid_cad = pt_to_cad(mid_pt)
                    if pts_contour:
                        s_snap = min(pts_contour, key=lambda q: math.hypot(q[0] - start_cad[0], q[1] - start_cad[1]))
                        e_snap = min(pts_contour, key=lambda q: math.hypot(q[0] - end_cad[0], q[1] - end_cad[1]))
                        if math.hypot(s_snap[0] - start_cad[0], s_snap[1] - start_cad[1]) <= 1.5:
                            start_cad = s_snap
                        if math.hypot(e_snap[0] - end_cad[0], e_snap[1] - end_cad[1]) <= 1.5:
                            end_cad = e_snap
                    snapped_arcs_data.append((start_cad, end_cad, c_cad, mid_cad))

            # 10. Extragere cicluri închise (LWPOLYLINE 100% continuu)
            closed_loops = self.extract_all_closed_loops(final_lines, snapped_arcs_data)

            # 11. Creare document DXF cu entități curate
            dxf_doc = ezdxf.new("R2010", setup=True)
            dxf_doc.units = ezdxf.units.MM

            # Configurare parametri de cotare metrică 1:1 (mm)
            # Forțăm ca stilul activ și toate stilurile de cotare să aibă factorul de scalare DIMLFAC = 1.0 (nu 100.0)
            dxf_doc.header["$DIMLFAC"] = 1.0
            dxf_doc.header["$DIMSCALE"] = 1.0
            dxf_doc.header["$DIMSTYLE"] = "Standard"
            dxf_doc.header["$MEASUREMENT"] = 1  # 1 = Metric
            dxf_doc.header["$INSUNITS"] = 4     # 4 = Millimeters
            dxf_doc.header["$LUNITS"] = 2       # 2 = Decimal
            dxf_doc.header["$DIMLUNIT"] = 2     # 2 = Decimal

            for ds in dxf_doc.dimstyles:
                ds.dxf.dimlfac = 1.0
                ds.dxf.dimscale = 1.0

            msp = dxf_doc.modelspace()

            if closed_loops:
                for loop in closed_loops:
                    msp.add_lwpolyline(
                        loop,
                        format="xyseb",
                        close=True,
                        dxfattribs={"layer": "0"}
                    )
            else:
                # Fallback dacă din motive excepționale nu se formează un loop complet închis
                for p1, p2 in final_lines:
                    msp.add_line(p1, p2, dxfattribs={"layer": "0"})

                for s_cad, e_cad, c_cad, m_cad in snapped_arcs_data:
                    r_cad = math.hypot(s_cad[0] - c_cad[0], s_cad[1] - c_cad[1])
                    a_start, a_end = self._get_arc_angles(c_cad[0], c_cad[1], s_cad, m_cad, e_cad)
                    msp.add_arc(
                        center=c_cad,
                        radius=r_cad,
                        start_angle=a_start,
                        end_angle=a_end,
                        dxfattribs={"layer": "0"},
                    )

            # 12. Adăugare găuri circulare interioare (CIRCLE pe layer 0)
            if holes:
                scale_for_holes = base_scale_factor if is_shortened else scale_final
                for h in holes:
                    h_center_pt = fitz.Point(h["center"][0], h["center"][1])
                    cad_center = pt_to_cad(h_center_pt)
                    cad_radius = h["radius"] * scale_for_holes
                    msp.add_circle(
                        center=cad_center,
                        radius=cad_radius,
                        dxfattribs={"layer": "0"}
                    )

            # 9. Salvare DXF / Conversie DWG
            os.makedirs(output_dir, exist_ok=True)
            dxf_filename = f"{clean_name}.dxf"
            dwg_filename = f"{clean_name}.dwg"
            dxf_path = os.path.join(output_dir, dxf_filename)
            dwg_path = os.path.join(output_dir, dwg_filename)

            try:
                dxf_doc.saveas(dxf_path)
            except PermissionError:
                fallback_dxf = os.path.join(output_dir, f"{clean_name}_nou.dxf")
                try:
                    dxf_doc.saveas(fallback_dxf)
                    dxf_path = fallback_dxf
                except Exception:
                    return False, None, f"Fișierul '{dxf_filename}' este deschis în alt program (ex. AutoCAD). Închideți fișierul și reîncercați."

            output_format_lower = str(output_format).lower().strip()
            saved_paths = []
            dwg_success = False

            if output_format_lower in ["dwg", "both"]:
                if self.is_odafc_available(self.oda_path):
                    try:
                        self._configure_odafc()
                        if os.path.exists(dwg_path):
                            try:
                                os.remove(dwg_path)
                            except Exception:
                                pass
                        odafc.export_dwg(dxf_doc, dwg_path, version="R2010")
                        if os.path.isfile(dwg_path) and os.path.getsize(dwg_path) > 0:
                            saved_paths.append(dwg_path)
                            dwg_success = True
                    except Exception:
                        dwg_success = False

                if not dwg_success:
                    saved_paths.append(dxf_path)

            if output_format_lower in ["dxf", "both"] or (output_format_lower == "dwg" and not dwg_success):
                if dxf_path not in saved_paths:
                    saved_paths.append(dxf_path)

            if output_format_lower == "dwg" and dwg_success and os.path.exists(dxf_path):
                try:
                    os.remove(dxf_path)
                except Exception:
                    pass

            primary_path = saved_paths[0] if saved_paths else dxf_path
            dim_info = f"{dim_x_final:.1f} x {dim_y_final:.1f} mm"

            if dwg_success:
                msg = f"Generat DWG nativ 1:1 ({dim_info}): {os.path.basename(dwg_path)}"
            else:
                msg = f"Generat DXF 1:1 ({dim_info}): {os.path.basename(dxf_path)}"

            return True, primary_path, msg

        except Exception as e:
            return False, None, f"Eroare procesare PDF: {e}"

    def convert_batch(
        self,
        pdf_paths,
        output_dir,
        output_format="dwg",
        only_front_view=True,
        calibrate_1to1=True,
        progress_callback=None,
        log_callback=None,
    ):
        """
        Convertește în masă o listă de fișiere PDF.
        """
        success_count = 0
        fail_count = 0
        results = []

        total = len(pdf_paths)
        for idx, pdf_p in enumerate(pdf_paths, 1):
            if log_callback:
                log_callback(f"[{idx}/{total}] Procesare: {os.path.basename(pdf_p)}...")

            ok, out_path, msg = self.convert_single_pdf(
                pdf_path=pdf_p,
                output_dir=output_dir,
                output_format=output_format,
                only_front_view=only_front_view,
                calibrate_1to1=calibrate_1to1,
            )

            if ok:
                success_count += 1
                if log_callback:
                    log_callback(f"  ✅ {msg}")
                results.append({"pdf": pdf_p, "out": out_path, "status": "OK"})
            else:
                fail_count += 1
                if log_callback:
                    log_callback(f"  ❌ {msg}")
                results.append({"pdf": pdf_p, "out": None, "status": "FAIL", "error": msg})

            if progress_callback:
                progress_callback(idx / total)

        return success_count, fail_count, results
