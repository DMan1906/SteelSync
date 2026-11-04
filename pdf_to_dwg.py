# -*- coding: utf-8 -*-
"""
pdf_to_dwg.py - Generator CAD (DWG / DXF) din Desene de Piesă PDF
Extrage doar geometria curată a piesei (1:1 în milimetri), fără cote, texte, săgeți de secțiune sau cartuș.
Suportă decupaje complexe, găuri și expansiunea exactă a vederilor scurtate (Tekla view shortening).
"""

import os
import re
import math
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
        "POS", "NO", "ANTALL", "PROFIL", "MATERIAL", "LENGDE", "VEKT", "LITTRA",
        "NORSKE", "BARESYSTEMER", "BÆRESYSTEMER", "SWECO", "FISKEBRYGGA", "STALPLATE",
        "STÅLPLATE", "RIB", "A4", "A3", "A2", "A1", "A0", "FORMAT", "DATO", "REV",
        "ENDRING", "UTF", "KONTR", "STATUS", "ANSVARLIG", "TITTEL", "OPPDRAGSGIVER",
        "OPPDRAGSNR", "TEGNINGSNUMMER", "TEGNINGSSTATUS", "ORIGINALFORMAT", "ARBEIDSTEGNING"
    }

    def __init__(self, oda_path=None):
        self.oda_path = oda_path
        self._configure_odafc()

    def _configure_odafc(self):
        """Configurează calea către ODA File Converter dacă este disponibilă."""
        if odafc is None:
            return
        if self.oda_path and os.path.isfile(self.oda_path):
            try:
                odafc.win_exec_path = self.oda_path
            except Exception:
                pass

    @staticmethod
    def is_available():
        """Verifică dacă bibliotecile necesare (PyMuPDF și ezdxf) sunt instalate."""
        return fitz is not None and ezdxf is not None

    def has_oda(self):
        return self.is_odafc_available(self.oda_path)

    @staticmethod
    def is_odafc_available(custom_path=None):
        """Verifică dacă ODA File Converter este configurat și funcțional pentru DWG nativ."""
        if odafc is None:
            return False
        if custom_path and os.path.isfile(custom_path):
            try:
                odafc.win_exec_path = custom_path
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
            "clean_name": None,
        }

        # 1. Determinare Y pentru STYKKLISTE / Cartuș
        for b in blocks:
            txt = b[4].lower()
            if any(k in txt for k in ["stykkliste", "arbeidstegning", "pos.nummer", "originalformat"]):
                if meta["stykk_y"] is None or b[1] < meta["stykk_y"]:
                    meta["stykk_y"] = b[1]

        # 2. Extragere din blocul de date STYKKLISTE
        for b in blocks:
            b_txt = b[4].strip()
            lines = [l.strip() for l in b_txt.splitlines() if l.strip()]
            for idx, line in enumerate(lines):
                # Detecție profil (ex: PL20X200, PL25X300, BL3*931)
                m_prof = re.search(r'\b(?:PL|BL|FL|P|FLAT)?\s*(\d+(?:\.\d+)?)\s*[*xX\-]\s*(\d+(?:\.\d+)?)\b', line, re.IGNORECASE)
                if m_prof:
                    try:
                        meta["profil_thickness"] = float(m_prof.group(1))
                        meta["profil_width"] = float(m_prof.group(2))
                    except ValueError:
                        pass

                    # Codul piesei este de regulă linia precedentă profilului
                    if idx > 0:
                        prev = lines[idx - 1].strip()
                        clean_prev = re.sub(r'[^A-Za-z0-9]', '', prev).upper()
                        if clean_prev and clean_prev not in self.BLACKLIST_NAMES and any(c.isdigit() for c in clean_prev):
                            meta["pos_no"] = prev

                    # Lungimea (LENGDE) este de regulă un număr pe una din liniile următoare
                    for sub in lines[idx + 1:]:
                        if sub.replace('.', '', 1).isdigit():
                            val = float(sub)
                            if val > 10 and (meta["lengde"] is None or val > meta["lengde"]):
                                meta["lengde"] = val
                                break
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

        # 6. Identificare Scară (Målestokk 1:X sau Front 1:X)
        m_scale = re.search(r'(?:M.*?lestokk|Scale|Front|Top|View)\s*[:\s]?\s*1\s*:\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
        if m_scale:
            try:
                meta["scale"] = float(m_scale.group(1))
            except ValueError:
                meta["scale"] = 10.0

        if not meta["lengde"]:
            m_len = re.search(r'LENGDE\s*[\r\n]+\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
            if m_len:
                try:
                    meta["lengde"] = float(m_len.group(1))
                except ValueError:
                    pass

        if not meta["profil_width"]:
            m_prof = re.search(r'\b(?:PL|BL|FL|P|FLAT)?\s*(\d+(?:\.\d+)?)\s*[*xX\-]\s*(\d+(?:\.\d+)?)\b', text, re.IGNORECASE)
            if m_prof:
                try:
                    meta["profil_thickness"] = float(m_prof.group(1))
                    meta["profil_width"] = float(m_prof.group(2))
                except ValueError:
                    pass

        return meta

    def extract_geometry_items(self, page, stykk_y=None, only_front_view=True):
        """
        Extrage geometria curată a piesei:
        - Filtrează cartușul și tabelele.
        - Elimină automat liniile de cotă, cotele, săgețile și simbolurile de secțiune (cutting planes).
        - Izolează vederea principală (Front) față de secțiunile laterale înguste.
        - Păstrează decupajele, găurile și colțurile 100% intacte (fără uniri accidentale prin decupaje).
        """
        max_y = (stykk_y - 8) if stykk_y else (page.rect.height * 0.72)
        margin_x0 = 20
        margin_x1 = page.rect.width - 20
        margin_y0 = 20

        drawings = page.get_drawings()
        graphic_drawings = [
            d for d in drawings
            if d["rect"].y1 < max_y and d["rect"].y0 > margin_y0 and d["rect"].x0 > margin_x0 and d["rect"].x1 < margin_x1
        ]

        if not graphic_drawings:
            return [], None

        # 1. Eliminare directă a simbolurilor de secțiune / cutting planes (w > 1.05 pt)
        non_section_drawings = [
            d for d in graphic_drawings
            if (d.get("width") or 0) <= 1.05
        ]

        if not non_section_drawings:
            non_section_drawings = graphic_drawings

        # 2. Identificare grosime pentru conturul obiectului (Visible outline)
        ns_widths = [round(d.get("width") or 0, 3) for d in non_section_drawings if d.get("width")]
        
        if any(0.65 <= w <= 0.85 for w in ns_widths):
            target_w_min, target_w_max = 0.65, 0.85
        elif any(0.50 <= w <= 0.65 for w in ns_widths):
            target_w_min, target_w_max = 0.50, 0.65
        elif ns_widths:
            max_ns = max(ns_widths)
            target_w_min, target_w_max = max_ns * 0.75, max_ns * 1.25
        else:
            target_w_min, target_w_max = 0.1, 1.0

        contour_drawings = [
            d for d in non_section_drawings
            if d.get("width") and target_w_min <= round(d.get("width"), 3) <= target_w_max and d.get("items")
        ]

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

        # 3.1 Izolare vedere principală pe axa Y (eliminare vederi secundare Top / Topp / Bunn)
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

            # Vederea Front are înălțimea conturului mult mai mare decât Topp (care este doar o bandă de grosime t)
            clusters_y.sort(key=lambda c: (c[1] - c[0]), reverse=True)
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
        bbox["width"] = bbox["max_x"] - bbox["min_x"]
        bbox["height"] = bbox["max_y"] - bbox["min_y"]

        return items, bbox

    @staticmethod
    def _find_view_break_pairs(items, min_x, max_x):
        """
        Detectează dacă vederea este scurtată de Tekla (view shortening break)
        și localizează perechile de segmente orizontale tăiate peste linia de ruptură.
        Returnează: (x_brk_left, x_brk_right, break_pairs, used_indices) sau None
        """
        w = max_x - min_x
        h_lines = []
        for idx, it in enumerate(items):
            if it[0] == "l" and abs(it[1].y - it[2].y) < 0.25:
                h_lines.append((idx, it))

        pairs = []
        used_indices = set()
        for i, (idx1, l1) in enumerate(h_lines):
            if idx1 in used_indices:
                continue
            p1a, p1b = l1[1], l1[2]
            l1_min_x, l1_max_x = min(p1a.x, p1b.x), max(p1a.x, p1b.x)
            l1_y = (p1a.y + p1b.y) / 2.0

            # Căutăm segmentul din stânga (capătul interior este în zona centrală)
            if min_x + 0.15 * w < l1_max_x < max_x - 0.15 * w:
                for j, (idx2, l2) in enumerate(h_lines):
                    if idx2 in used_indices or idx1 == idx2:
                        continue
                    p2a, p2b = l2[1], l2[2]
                    l2_min_x, l2_max_x = min(p2a.x, p2b.x), max(p2a.x, p2b.x)
                    l2_y = (p2a.y + p2b.y) / 2.0

                    if abs(l1_y - l2_y) < 0.3:
                        gap = l2_min_x - l1_max_x
                        if 3.0 <= gap <= 40.0:
                            outer_left = p1a if p1a.x < p1b.x else p1b
                            inner_left = p1b if p1a.x < p1b.x else p1a
                            inner_right = p2a if p2a.x < p2b.x else p2b
                            outer_right = p2b if p2a.x < p2b.x else p2a
                            pairs.append((idx1, idx2, outer_left, inner_left, inner_right, outer_right))
                            used_indices.add(idx1)
                            used_indices.add(idx2)
                            break

        if not pairs:
            return None

        x_brk_left = max(p[3].x for p in pairs)
        x_brk_right = min(p[4].x for p in pairs)
        return x_brk_left, x_brk_right, pairs, used_indices

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

    @classmethod
    def _detect_fillet_arcs(cls, items, exclude_indices=None):
        """
        Identifică arcele de cerc (fillet) reprezentate în PDF prin secvențe
        de micro-linii tangente și curbe Bezier cubice (l + c + l sau doar c).
        Găsește centrul real, raza exactă și unghiurile arcului.
        Returnează: (fillet_arcs, fillet_used_indices)
        unde fiecare element din fillet_arcs este (cx_pdf, cy_pdf, r_pdf, start_pt, mid_pt, end_pt).
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

            # Căutăm linii mici tangente adiacente la capetele curbei (l <= 1.2 pt)
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

            # Punct median pe curba Bezier la t=0.5
            t = 0.5
            bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
            by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
            mid_pt = fitz.Point(bx, by)

            fit = cls._fit_circle_3pts((start_pt.x, start_pt.y), (mid_pt.x, mid_pt.y), (end_pt.x, end_pt.y))
            if fit:
                cx_pdf, cy_pdf, r_pdf = fit
                # Verificăm dacă punctele de control se potrivesc cu cercul cu o eroare < 5%
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
                page, stykk_y=meta["stykk_y"], only_front_view=only_front_view
            )
            doc.close()

            if not items or not bbox or bbox["width"] <= 0 or bbox["height"] <= 0:
                return False, None, "Nu s-a putut izola conturul piesei (lipsă vectori sau desen raster)."

            base_scale_factor = (25.4 / 72.0) * meta["scale"]
            min_x = bbox["min_x"]
            max_x = bbox["max_x"]
            min_y = bbox["min_y"]
            max_y = bbox["max_y"]

            drawn_w_pt = bbox["width"]
            drawn_h_pt = bbox["height"]
            drawn_len_mm = drawn_w_pt * base_scale_factor

            target_len = meta["lengde"]
            target_wid = meta["profil_width"]

            # Verificare dacă este o vedere scurtată grafic (Tekla view shortening)
            view_break_info = None
            if calibrate_1to1 and target_len and (target_len / drawn_len_mm >= 1.35):
                view_break_info = self._find_view_break_pairs(items, min_x, max_x)

            # Identificare arce de racordare (fillet) pentru a le exporta ca arce native de cerc DXF (ARC)
            fillet_arcs, fillet_used_indices = self._detect_fillet_arcs(
                items, exclude_indices=view_break_info[3] if view_break_info else None
            )

            dxf_doc = ezdxf.new("R2010", setup=True)
            dxf_doc.units = ezdxf.units.MM
            msp = dxf_doc.modelspace()
            dxf_doc.layers.add("CONTUR_PIESA", color=4)  # Cyan

            if view_break_info is not None:
                # ── Caz 1: Vedere Scurtată (Tekla View Shortening) ──
                # Unim direct segmentele tăiate într-o singură linie continuă din capăt în capăt,
                # eliminând golurile, decuplările și liniile fragmentate.
                x_brk_left, x_brk_right, break_pairs, used_indices = view_break_info

                left_mm = (x_brk_left - min_x) * base_scale_factor
                right_mm = (max_x - x_brk_right) * base_scale_factor
                delta_gap_mm = target_len - (left_mm + right_mm)

                def pt_to_cad(pt):
                    if pt.x <= x_brk_left + 0.5:
                        xc = (pt.x - min_x) * base_scale_factor
                    else:
                        xc = left_mm + delta_gap_mm + (pt.x - x_brk_right) * base_scale_factor
                    yc = (max_y - pt.y) * base_scale_factor
                    return (float(xc), float(yc))

                # 1. Conectăm direct capetele exterioare ale fiecărei muchii tăiate într-o singură linie continuă
                for pair in break_pairs:
                    outer_left, outer_right = pair[2], pair[5]
                    p_start = pt_to_cad(outer_left)
                    p_end = pt_to_cad(outer_right)
                    msp.add_line(p_start, p_end, dxfattribs={"layer": "CONTUR_PIESA"})

                # 2. Adăugăm arcele de racordare (fillet) ca arce native de cerc DXF (ARC)
                for cx_pdf, cy_pdf, r_pdf, start_pt, mid_pt, end_pt in fillet_arcs:
                    c_cad = pt_to_cad(fitz.Point(cx_pdf, cy_pdf))
                    start_cad = pt_to_cad(start_pt)
                    end_cad = pt_to_cad(end_pt)
                    mid_cad = pt_to_cad(mid_pt)
                    r_cad = math.hypot(start_cad[0] - c_cad[0], start_cad[1] - c_cad[1])
                    a_start, a_end = self._get_arc_angles(c_cad[0], c_cad[1], start_cad, mid_cad, end_cad)
                    msp.add_arc(
                        center=c_cad,
                        radius=r_cad,
                        start_angle=a_start,
                        end_angle=a_end,
                        dxfattribs={"layer": "CONTUR_PIESA"},
                    )

                # 3. Exportăm restul elementelor geometrice intacte (decupaje, colțuri, găuri, raze)
                all_skipped = used_indices | fillet_used_indices
                for idx, it in enumerate(items):
                    if idx in all_skipped:
                        continue  # Muchiile longitudinale și arcele fillet au fost deja create

                    cmd = it[0]
                    if cmd == "l":
                        p1 = pt_to_cad(it[1])
                        p2 = pt_to_cad(it[2])
                        msp.add_line(p1, p2, dxfattribs={"layer": "CONTUR_PIESA"})

                    elif cmd == "c":
                        p1, p2, p3, p4 = it[1], it[2], it[3], it[4]
                        steps = 16
                        pts = []
                        for s in range(steps + 1):
                            t = s / float(steps)
                            bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
                            by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
                            pts.append(pt_to_cad(fitz.Point(bx, by)))
                        msp.add_lwpolyline(pts, dxfattribs={"layer": "CONTUR_PIESA"})

                    elif cmd == "re":
                        r = it[1]
                        p0 = pt_to_cad(fitz.Point(r.x0, r.y0))
                        p1 = pt_to_cad(fitz.Point(r.x1, r.y0))
                        p2 = pt_to_cad(fitz.Point(r.x1, r.y1))
                        p3 = pt_to_cad(fitz.Point(r.x0, r.y1))
                        msp.add_lwpolyline([p0, p1, p2, p3, p0], dxfattribs={"layer": "CONTUR_PIESA"})

                dim_x_final = target_len
                dim_y_final = drawn_h_pt * base_scale_factor

            else:
                # ── Caz 2: Vedere Normală (Nescurtată) ──
                scale_final = base_scale_factor
                if calibrate_1to1 and target_len and target_wid:
                    target_max = max(target_len, target_wid)
                    curr_max_mm = max(drawn_w_pt, drawn_h_pt) * base_scale_factor
                    if curr_max_mm > 0:
                        ratio = target_max / curr_max_mm
                        if 0.85 <= ratio <= 1.15:
                            scale_final = base_scale_factor * ratio

                dim_x_final = drawn_w_pt * scale_final
                dim_y_final = drawn_h_pt * scale_final

                # 1. Adăugăm arcele de racordare (fillet) ca arce native de cerc DXF (ARC)
                for cx_pdf, cy_pdf, r_pdf, start_pt, mid_pt, end_pt in fillet_arcs:
                    cx_cad = (cx_pdf - min_x) * scale_final
                    cy_cad = (max_y - cy_pdf) * scale_final
                    sx_cad = (start_pt.x - min_x) * scale_final
                    sy_cad = (max_y - start_pt.y) * scale_final
                    ex_cad = (end_pt.x - min_x) * scale_final
                    ey_cad = (max_y - end_pt.y) * scale_final
                    mx_cad = (mid_pt.x - min_x) * scale_final
                    my_cad = (max_y - mid_pt.y) * scale_final
                    r_cad = math.hypot(sx_cad - cx_cad, sy_cad - cy_cad)
                    a_start, a_end = self._get_arc_angles(cx_cad, cy_cad, (sx_cad, sy_cad), (mx_cad, my_cad), (ex_cad, ey_cad))
                    msp.add_arc(
                        center=(cx_cad, cy_cad),
                        radius=r_cad,
                        start_angle=a_start,
                        end_angle=a_end,
                        dxfattribs={"layer": "CONTUR_PIESA"},
                    )

                # 2. Exportăm restul elementelor geometrice intacte
                for idx, it in enumerate(items):
                    if idx in fillet_used_indices:
                        continue

                    cmd = it[0]
                    if cmd == "l":
                        p1, p2 = it[1], it[2]
                        x1 = (p1.x - min_x) * scale_final
                        y1 = (max_y - p1.y) * scale_final
                        x2 = (p2.x - min_x) * scale_final
                        y2 = (max_y - p2.y) * scale_final
                        msp.add_line((x1, y1), (x2, y2), dxfattribs={"layer": "CONTUR_PIESA"})

                    elif cmd == "c":
                        p1, p2, p3, p4 = it[1], it[2], it[3], it[4]
                        steps = 16
                        pts = []
                        for s in range(steps + 1):
                            t = s / float(steps)
                            bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
                            by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
                            cx = (bx - min_x) * scale_final
                            cy = (max_y - by) * scale_final
                            pts.append((cx, cy))
                        msp.add_lwpolyline(pts, dxfattribs={"layer": "CONTUR_PIESA"})

                    elif cmd == "re":
                        r = it[1]
                        rx0 = (r.x0 - min_x) * scale_final
                        rx1 = (r.x1 - min_x) * scale_final
                        ry0 = (max_y - r.y1) * scale_final
                        ry1 = (max_y - r.y0) * scale_final
                        poly_pts = [(rx0, ry0), (rx1, ry0), (rx1, ry1), (rx0, ry1), (rx0, ry0)]
                        msp.add_lwpolyline(poly_pts, dxfattribs={"layer": "CONTUR_PIESA"})

            # 4. Salvare DXF / DWG
            os.makedirs(output_dir, exist_ok=True)
            dxf_filename = f"{clean_name}.dxf"
            dwg_filename = f"{clean_name}.dwg"
            dxf_path = os.path.join(output_dir, dxf_filename)
            dwg_path = os.path.join(output_dir, dwg_filename)

            try:
                dxf_doc.saveas(dxf_path)
            except PermissionError:
                # Dacă fișierul original este deschis în AutoCAD / alt program, salvăm temporar cu sufix _nou
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
