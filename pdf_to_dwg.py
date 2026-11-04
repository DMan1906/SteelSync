# -*- coding: utf-8 -*-
"""
pdf_to_dwg.py - Generator CAD (DWG / DXF) din Desene de Piesă PDF
Extrage doar geometria curată a piesei (1:1 în milimetri), fără cote, texte, săgeți de secțiune sau cartuș.
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
        """Curăță denumirea piesei de cratime, spații și caractere speciale (ex: DSBO-21159 -> DSBO21159)."""
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
                # Detecție profil (ex: PL25X300, BL3*931, FL100X10, PL20-294)
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
        # Dacă numele fișierului PDF este un cod alfanumeric valid (ex: DSBO21159_.pdf), îl folosim ca garanție
        if clean_stem and clean_stem not in self.BLACKLIST_NAMES and any(c.isdigit() for c in clean_stem):
            if not meta["pos_no"] or meta["pos_no"].upper() in self.BLACKLIST_NAMES:
                meta["pos_no"] = clean_stem
            else:
                # Dacă codul extras e aproape identic cu stem-ul, preferăm forma curată a numelui
                pass
        elif not meta["pos_no"]:
            meta["pos_no"] = stem

        meta["clean_name"] = self.clean_part_name(meta["pos_no"])

        # 6. Identificare Scară (Målestokk 1:X sau Front 1:X)
        m_scale = re.search(r'(?:M[aå]lestokk|Scale|Front|Top|View)\s*[:\s]\s*1\s*:\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
        if m_scale:
            try:
                meta["scale"] = float(m_scale.group(1))
            except ValueError:
                meta["scale"] = 10.0

        # Fallback LENGDE / PROFIL din text general dacă nu au fost găsite în blocul STYKKLISTE
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
        - Unește segmentele coliniare întrerupte de simboluri de secțiune pentru a obține conturul întreg.
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

        # 1. Eliminare directă a liniilor de secțiune (cutting planes)
        # În standardele Tekla/CAD, liniile de secțiune / simbolurile de secțiune sunt trasate cu Pen gros (w > 1.05 pt)
        non_section_drawings = [
            d for d in graphic_drawings
            if (d.get("width") or 0) <= 1.05
        ]

        if not non_section_drawings:
            non_section_drawings = graphic_drawings

        # 2. Identificare grosime pentru conturul obiectului (Visible outline)
        # În Tekla metric, conturul este de regulă ~0.709 pt (sau ~0.567 pt)
        # Liniile de cotă / hașurile sunt subțiri (~0.28 - 0.51 pt)
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
            # Fallback dacă nu sunt diferențe clare de grosime
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

            # Selectăm clusterul cel mai lat (Front view are lățimea piesei, în timp ce secțiunea A-A are doar grosimea t)
            clusters.sort(key=lambda c: (c[1] - c[0]), reverse=True)
            contour_drawings = clusters[0][2]

        # 4. Colectare primitive geometrice
        raw_items = []
        for d in contour_drawings:
            for it in d.get("items", []):
                raw_items.append(it)

        if not raw_items:
            return [], None

        # 5. Unire segmente coliniare întrerupte de simboluri de secțiune sau axe
        items = self._merge_collinear_lines(raw_items, max_gap=35.0)

        # 6. Calcul Bounding Box
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
    def _merge_collinear_lines(items, max_gap=35.0):
        """
        Unește segmentele de linie coliniare întrerupte de goluri mici
        (generate de trecerea liniilor de secțiune sau axe prin conturul piesei).
        """
        lines = [it for it in items if it[0] == "l"]
        non_lines = [it for it in items if it[0] != "l"]

        if not lines:
            return items

        h_lines = []
        v_lines = []
        other_lines = []

        for it in lines:
            p1, p2 = it[1], it[2]
            if abs(p1.y - p2.y) < 0.6:
                y = (p1.y + p2.y) / 2.0
                x_min, x_max = min(p1.x, p2.x), max(p1.x, p2.x)
                h_lines.append((y, x_min, x_max))
            elif abs(p1.x - p2.x) < 0.6:
                x = (p1.x + p2.x) / 2.0
                y_min, y_max = min(p1.y, p2.y), max(p1.y, p2.y)
                v_lines.append((x, y_min, y_max))
            else:
                other_lines.append(it)

        merged = []

        # Unire orizontale
        h_by_y = defaultdict(list)
        for y, x0, x1 in h_lines:
            matched_y = None
            for ey in h_by_y:
                if abs(ey - y) < 0.6:
                    matched_y = ey
                    break
            if matched_y is None:
                matched_y = y
            h_by_y[matched_y].append((x0, x1))

        for y, spans in h_by_y.items():
            spans.sort()
            cur_x0, cur_x1 = spans[0]
            for s0, s1 in spans[1:]:
                if s0 - cur_x1 <= max_gap:
                    cur_x1 = max(cur_x1, s1)
                else:
                    merged.append(("l", fitz.Point(cur_x0, y), fitz.Point(cur_x1, y)))
                    cur_x0, cur_x1 = s0, s1
            merged.append(("l", fitz.Point(cur_x0, y), fitz.Point(cur_x1, y)))

        # Unire verticale
        v_by_x = defaultdict(list)
        for x, y0, y1 in v_lines:
            matched_x = None
            for ex in v_by_x:
                if abs(ex - x) < 0.6:
                    matched_x = ex
                    break
            if matched_x is None:
                matched_x = x
            v_by_x[matched_x].append((y0, y1))

        for x, spans in v_by_x.items():
            spans.sort()
            cur_y0, cur_y1 = spans[0]
            for s0, s1 in spans[1:]:
                if s0 - cur_y1 <= max_gap:
                    cur_y1 = max(cur_y1, s1)
                else:
                    merged.append(("l", fitz.Point(x, cur_y0), fitz.Point(x, cur_y1)))
                    cur_y0, cur_y1 = s0, s1
            merged.append(("l", fitz.Point(x, cur_y0), fitz.Point(x, cur_y1)))

        merged.extend(other_lines)
        merged.extend(non_lines)
        return merged

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
        output_format: 'dwg', 'dxf', sau 'both'
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

            # 3. Scalare și Calibrare 1:1
            # 1 pt = 25.4 / 72.0 mm
            base_scale_factor = (25.4 / 72.0) * meta["scale"]
            scale_x = base_scale_factor
            scale_y = base_scale_factor

            # Calibrare precisă când avem LENGDE și PROFIL
            if calibrate_1to1 and meta["lengde"] and meta["profil_width"]:
                target_len = meta["lengde"]
                target_wid = meta["profil_width"]

                # Verificăm orientarea vederii desenate (orizontală vs verticală)
                if bbox["width"] >= bbox["height"]:
                    target_x = target_len
                    target_y = target_wid
                else:
                    target_x = target_wid
                    target_y = target_len

                curr_x_mm = bbox["width"] * base_scale_factor
                curr_y_mm = bbox["height"] * base_scale_factor

                ratio_x = target_x / curr_x_mm if curr_x_mm > 0 else 1.0
                ratio_y = target_y / curr_y_mm if curr_y_mm > 0 else 1.0

                # Dacă piesa a fost desenată scurtată pe lungime (Tekla view shortening),
                # calibrăm axa lungimii exact la valoarea nominală din STYKKLISTE
                if 0.5 <= ratio_y <= 2.0:
                    scale_y = base_scale_factor * ratio_y
                if ratio_x > 0:
                    scale_x = base_scale_factor * ratio_x

            dim_x_final = bbox["width"] * scale_x
            dim_y_final = bbox["height"] * scale_y

            # 4. Construire document CAD cu ezdxf
            dxf_doc = ezdxf.new("R2010", setup=True)
            dxf_doc.units = ezdxf.units.MM
            msp = dxf_doc.modelspace()

            min_x = bbox["min_x"]
            max_y_pt = bbox["max_y"]

            dxf_doc.layers.add("CONTUR_PIESA", color=4)  # Cyan

            for it in items:
                cmd = it[0]
                if cmd == "l":
                    p1, p2 = it[1], it[2]
                    x1 = (p1.x - min_x) * scale_x
                    y1 = (max_y_pt - p1.y) * scale_y
                    x2 = (p2.x - min_x) * scale_x
                    y2 = (max_y_pt - p2.y) * scale_y
                    msp.add_line((x1, y1), (x2, y2), dxfattribs={"layer": "CONTUR_PIESA"})

                elif cmd == "c":
                    # Interpolare Bezier cubică
                    p1, p2, p3, p4 = it[1], it[2], it[3], it[4]
                    steps = 16
                    pts = []
                    for s in range(steps + 1):
                        t = s / float(steps)
                        bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
                        by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
                        cx = (bx - min_x) * scale_x
                        cy = (max_y_pt - by) * scale_y
                        pts.append((cx, cy))
                    msp.add_lwpolyline(pts, dxfattribs={"layer": "CONTUR_PIESA"})

                elif cmd == "re":
                    r = it[1]
                    rx0 = (r.x0 - min_x) * scale_x
                    rx1 = (r.x1 - min_x) * scale_x
                    ry0 = (max_y_pt - r.y1) * scale_y
                    ry1 = (max_y_pt - r.y0) * scale_y
                    poly_pts = [(rx0, ry0), (rx1, ry0), (rx1, ry1), (rx0, ry1), (rx0, ry0)]
                    msp.add_lwpolyline(poly_pts, dxfattribs={"layer": "CONTUR_PIESA"})

            # 5. Salvare DXF / DWG
            os.makedirs(output_dir, exist_ok=True)
            dxf_filename = f"{clean_name}.dxf"
            dwg_filename = f"{clean_name}.dwg"
            dxf_path = os.path.join(output_dir, dxf_filename)
            dwg_path = os.path.join(output_dir, dwg_filename)

            dxf_doc.saveas(dxf_path)

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
            name_base = Path(pdf_p).stem
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
