# -*- coding: utf-8 -*-
"""
pdf_to_dwg.py - Generator CAD (DWG / DXF) din Desene de Piesă PDF
Extrage doar geometria curată a piesei (1:1 în milimetri), fără cote, texte, săgeți sau cartuș.
"""

import os
import re
import math
import shutil
from pathlib import Path

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

    @staticmethod
    def clean_part_name(raw_name):
        """Curăță denumirea piesei de cratime, spații și caractere speciale (ex: DSBO-21021 -> DSBO21021)."""
        if not raw_name:
            return "PART"
        clean = re.sub(r'[^A-Za-z0-9]', '', raw_name).upper()
        return clean if clean else "PART"

    def extract_drawing_metadata(self, page, pdf_path):
        """
        Extrage metadatele din textul paginii:
        - Codul piesei (POS. NO. / POS. NUMMER)
        - Scara desenului (Målestokk 1:X, Front 1:X)
        - Dimensiunile nominale (LENGDE, PROFIL)
        - Coordonata Y de început a tabelului STYKKLISTE
        """
        text = page.get_text()
        words = page.get_text("words")  # (x0, y0, x1, y1, word, block_no, line_no, word_no)
        
        meta = {
            "pos_no": None,
            "scale": 10.0,
            "lengde": None,
            "profil_width": None,
            "profil_thickness": None,
            "stykk_y": None,
            "clean_name": None,
        }

        # 1. Identificare Y pentru STYKKLISTE / Cartuș
        for w in words:
            w_text = w[4].lower()
            if any(k in w_text for k in ["stykkliste", "arbeidstegning", "pos.nummer", "originalformat"]):
                if meta["stykk_y"] is None or w[1] < meta["stykk_y"]:
                    meta["stykk_y"] = w[1]

        # 2. Identificare Cod Piesă (POS. NO.)
        # Căutare sub header POS. NO. în tabel
        pos_hdr = [w for w in words if w[4].upper() in ["POS.", "POS.NO", "POS.NO."]]
        if pos_hdr:
            hx0, hy0, hx1, hy1, _ = pos_hdr[0][:5]
            below = [w for w in words if abs(w[0] - hx0) < 35 and hy1 < w[1] < hy1 + 45]
            if below:
                candidate = below[0][4].strip()
                if candidate.upper() not in ["POS", "NO", "ANTALL", "PROFIL", "MATERIAL", "LENGDE", "VEKT"]:
                    meta["pos_no"] = candidate

        # Căutare în 'STYKKLISTE FOR <COD>'
        if not meta["pos_no"]:
            m_for = re.search(r'STYKKLISTE\s+FOR\s+([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
            if m_for:
                candidate = m_for.group(1).strip()
                if candidate.upper() not in ["POS", "NO", "ANTALL", "PROFIL", "MATERIAL", "LENGDE", "VEKT"]:
                    meta["pos_no"] = candidate

        # Căutare în cartuș după Tegn.nr / Pos. nummer
        if not meta["pos_no"]:
            m_teg = re.search(r'(?:Pos\.\s*nummer|Tegn\.nr\.?)\s*[\r\n]+\s*([A-Za-z0-9_\-]+)', text, re.IGNORECASE)
            if m_teg:
                meta["pos_no"] = m_teg.group(1).strip()

        # Fallback: Numele fișierului PDF
        if not meta["pos_no"]:
            stem = Path(pdf_path).stem
            meta["pos_no"] = stem

        meta["clean_name"] = self.clean_part_name(meta["pos_no"])

        # 3. Identificare Scară (Målestokk 1:X sau Front 1:X)
        m_scale = re.search(r'(?:M[aå]lestokk|Scale|Front|Top|View)\s*[:\s]\s*1\s*:\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
        if m_scale:
            try:
                meta["scale"] = float(m_scale.group(1))
            except ValueError:
                meta["scale"] = 10.0

        # 4. Identificare LENGDE
        m_len = re.search(r'LENGDE\s*[\r\n]+\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
        if m_len:
            try:
                meta["lengde"] = float(m_len.group(1))
            except ValueError:
                pass

        # 5. Identificare PROFIL (ex: PL20X294, BL3*931, FL100X10, 6x124)
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
        Extrage segmentele vectoriale aparținând piesei:
        - Filtrează căile din afara ariei grafice (cartuș, tabele).
        - Filtrează liniile subțiri de cotă și hașuri (păstrând doar conturul piesei).
        - Convertește căile în primitive geometrice (linii, arce/beziere, dreptunghiuri).
        """
        max_y = (stykk_y - 8) if stykk_y else (page.rect.height * 0.72)
        margin_x0 = 20
        margin_x1 = page.rect.width - 20
        margin_y0 = 20

        drawings = page.get_drawings()
        # Filtrare preliminară în spațiul grafic
        graphic_drawings = [
            d for d in drawings
            if d["rect"].y1 < max_y and d["rect"].y0 > margin_y0 and d["rect"].x0 > margin_x0 and d["rect"].x1 < margin_x1
        ]

        if not graphic_drawings:
            return [], None

        # Determinare prag de grosime (stroke width) pentru conturul real al piesei
        widths = [round(d.get("width") or 0, 3) for d in graphic_drawings if d.get("width")]
        if not widths:
            max_w = 0.5
        else:
            max_w = max(widths)

        # Conturul piesei are grosime considerabil mai mare decât liniile de cotă
        thick_threshold = max_w * 0.72

        # Păstrăm desenele cu linie groasă și eliminăm elementele de tip fill-only fără stroke
        contour_drawings = [
            d for d in graphic_drawings
            if (d.get("width") or 0) >= thick_threshold and d.get("items")
        ]

        if not contour_drawings:
            # Fallback: dacă nu sunt variații de grosime, elimină doar umplerile pure
            contour_drawings = [d for d in graphic_drawings if d.get("width") and d.get("items")]

        # Extragere primitive geometrice
        raw_items = []
        for d in contour_drawings:
            for it in d["items"]:
                raw_items.append(it)

        if not raw_items:
            return [], None

        # Colectare puncte pentru determinare bounding box
        pts = []
        for it in raw_items:
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

        # Dacă se dorește doar vederea principală (Front) și există vederi multiple (Front + Secțiune laterală)
        if only_front_view and bbox["width"] > 50:
            raw_items, bbox = self._isolate_main_view(raw_items, bbox)

        return raw_items, bbox

    def _isolate_main_view(self, items, total_bbox):
        """Izolează clusterul vederii principale dacă există o secțiune de grosime separată pe desen."""
        # Grupăm elementele în funcție de poziția orizontală X
        # În desenele Tekla, vederea 'Front' este clusterul mare din stânga/centru
        # iar secțiunea 'A - A' este o bandă îngustă verticală în dreapta
        mid_x = (total_bbox["min_x"] + total_bbox["max_x"]) / 2

        # Calculăm centrele X ale fiecărui item
        item_centers = []
        for it in items:
            cmd = it[0]
            if cmd == "l":
                cx = (it[1].x + it[2].x) / 2
            elif cmd == "c":
                cx = (it[1].x + it[4].x) / 2
            elif cmd == "re":
                cx = (it[1].x0 + it[1].x1) / 2
            else:
                cx = total_bbox["min_x"]
            item_centers.append(cx)

        # Verificăm dacă există o separare clară în două clustere pe X
        gap_threshold = 25.0  # spațiu liber între vederi
        left_items = [it for it, cx in zip(items, item_centers) if cx < mid_x]
        right_items = [it for it, cx in zip(items, item_centers) if cx >= mid_x]

        if left_items and right_items:
            max_left_x = max(
                it[1].x if it[0] == "l" else (it[1].x0 if it[0] == "re" else it[1].x)
                for it in left_items
            )
            min_right_x = min(
                it[1].x if it[0] == "l" else (it[1].x0 if it[0] == "re" else it[1].x)
                for it in right_items
            )
            if min_right_x - max_left_x > gap_threshold:
                # Vederea principală este clusterul cu mai multe elemente sau arie mai mare
                main_items = left_items if len(left_items) >= len(right_items) else right_items
                pts = []
                for it in main_items:
                    cmd = it[0]
                    if cmd == "l":
                        pts.extend([it[1], it[2]])
                    elif cmd == "c":
                        pts.extend([it[1], it[2], it[3], it[4]])
                    elif cmd == "re":
                        r = it[1]
                        pts.extend([fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y1)])
                new_bbox = {
                    "min_x": min(p.x for p in pts),
                    "max_x": max(p.x for p in pts),
                    "min_y": min(p.y for p in pts),
                    "max_y": max(p.y for p in pts),
                }
                new_bbox["width"] = new_bbox["max_x"] - new_bbox["min_x"]
                new_bbox["height"] = new_bbox["max_y"] - new_bbox["min_y"]
                return main_items, new_bbox

        return items, total_bbox

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
        Convertește un singur PDF într-un fișier CAD (.dwg sau .dxf).
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

            # 3. Calcul scală și calibrare 1:1
            # Relație PDF points -> mm: 1 pt = 25.4 / 72.0 mm
            base_factor = (25.4 / 72.0) * meta["scale"]
            scale_final = base_factor

            w_mm_base = bbox["width"] * base_factor
            h_mm_base = bbox["height"] * base_factor

            # Calibrare exactă pe baza dimensiunilor din STYKKLISTE
            if calibrate_1to1 and meta["lengde"] and meta["profil_width"]:
                target_max = max(meta["lengde"], meta["profil_width"])
                curr_max = max(w_mm_base, h_mm_base)
                if curr_max > 0:
                    ratio = target_max / curr_max
                    # Dacă raportul este într-un interval rezonabil de calibrare (0.75 - 1.25), calibrează
                    if 0.75 <= ratio <= 1.25:
                        scale_final = base_factor * ratio

            dim_x_final = bbox["width"] * scale_final
            dim_y_final = bbox["height"] * scale_final

            # 4. Construire document CAD cu ezdxf
            dxf_doc = ezdxf.new("R2010", setup=True)
            dxf_doc.units = ezdxf.units.MM
            msp = dxf_doc.modelspace()

            min_x = bbox["min_x"]
            max_y_pt = bbox["max_y"]

            # Adăugare layer dedicat
            dxf_doc.layers.add("CONTUR_PIESA", color=4)  # Cyan

            for it in items:
                cmd = it[0]
                if cmd == "l":
                    p1, p2 = it[1], it[2]
                    x1 = (p1.x - min_x) * scale_final
                    y1 = (max_y_pt - p1.y) * scale_final
                    x2 = (p2.x - min_x) * scale_final
                    y2 = (max_y_pt - p2.y) * scale_final
                    msp.add_line((x1, y1), (x2, y2), dxfattribs={"layer": "CONTUR_PIESA"})

                elif cmd == "c":
                    # Interpolare curbă Bezier cubică în segmente line / polilinie fină
                    p1, p2, p3, p4 = it[1], it[2], it[3], it[4]
                    steps = 16
                    pts = []
                    for s in range(steps + 1):
                        t = s / float(steps)
                        bx = (1 - t)**3 * p1.x + 3 * (1 - t)**2 * t * p2.x + 3 * (1 - t) * t**2 * p3.x + t**3 * p4.x
                        by = (1 - t)**3 * p1.y + 3 * (1 - t)**2 * t * p2.y + 3 * (1 - t) * t**2 * p3.y + t**3 * p4.y
                        cx = (bx - min_x) * scale_final
                        cy = (max_y_pt - by) * scale_final
                        pts.append((cx, cy))
                    msp.add_lwpolyline(pts, dxfattribs={"layer": "CONTUR_PIESA"})

                elif cmd == "re":
                    r = it[1]
                    rx0 = (r.x0 - min_x) * scale_final
                    rx1 = (r.x1 - min_x) * scale_final
                    ry0 = (max_y_pt - r.y1) * scale_final
                    ry1 = (max_y_pt - r.y0) * scale_final
                    poly_pts = [(rx0, ry0), (rx1, ry0), (rx1, ry1), (rx0, ry1), (rx0, ry0)]
                    msp.add_lwpolyline(poly_pts, dxfattribs={"layer": "CONTUR_PIESA"})

            # 5. Salvare DXF / DWG
            os.makedirs(output_dir, exist_ok=True)
            dxf_filename = f"{clean_name}.dxf"
            dwg_filename = f"{clean_name}.dwg"
            dxf_path = os.path.join(output_dir, dxf_filename)
            dwg_path = os.path.join(output_dir, dwg_filename)

            # Salvare DXF bază
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
                    # Fallback automat la DXF dacă ODA File Converter nu este instalat
                    saved_paths.append(dxf_path)

            if output_format_lower in ["dxf", "both"] or (output_format_lower == "dwg" and not dwg_success):
                if dxf_path not in saved_paths:
                    saved_paths.append(dxf_path)

            # Dacă utilizatorul a cerut doar DWG și conversia a reușit, putem șterge temporarul DXF
            if output_format_lower == "dwg" and dwg_success and os.path.exists(dxf_path):
                try:
                    os.remove(dxf_path)
                except Exception:
                    pass

            primary_path = saved_paths[0] if saved_paths else dxf_path
            dim_info = f"{dim_x_final:.1f} x {dim_y_final:.1f} mm"

            if dwg_success:
                msg = f"Generat cu succes: {os.path.basename(dwg_path)} ({dim_info})"
            elif output_format_lower == "dwg":
                msg = f"Generat DXF 1:1: {os.path.basename(dxf_path)} ({dim_info}) — Notă: ODA Converter nu a fost detectat pentru .dwg nativ."
            else:
                msg = f"Generat cu succes: {os.path.basename(dxf_path)} ({dim_info})"

            return True, primary_path, msg

        except Exception as e:
            return False, None, f"Eroare procesare PDF: {str(e)}"

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
        Returnează (success_count, fail_count, results_list).
        """
        success_count = 0
        fail_count = 0
        results = []
        total = len(pdf_paths)

        for idx, pdf in enumerate(pdf_paths, start=1):
            if log_callback:
                log_callback(f"[{idx}/{total}] Procesare: {os.path.basename(pdf)}...")

            ok, out_path, msg = self.convert_single_pdf(
                pdf_path=pdf,
                output_dir=output_dir,
                output_format=output_format,
                only_front_view=only_front_view,
                calibrate_1to1=calibrate_1to1,
            )

            if ok:
                success_count += 1
                results.append({"pdf": pdf, "status": "OK", "output": out_path, "message": msg})
                if log_callback:
                    log_callback(f"  ✔ {msg}")
            else:
                fail_count += 1
                results.append({"pdf": pdf, "status": "ERROR", "output": None, "message": msg})
                if log_callback:
                    log_callback(f"  ❌ {os.path.basename(pdf)}: {msg}")

            if progress_callback and total > 0:
                progress_callback(idx / total)

        return success_count, fail_count, results

