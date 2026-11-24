# -*- coding: utf-8 -*-
"""
cheson_extractor.py - Motor DSTV (.nc1) + BOM pentru descompunerea profilelor compuse
(Tip 1: Cheson cu Talpă, Tip 2: Cheson fără Talpă, Tip 3: Cheson Dublu) cu blocare exactă
a cotelor de gabarit (±0.001 mm) și generare fișiere individuale DXF/DWG pe grosimi.

Construit exclusiv pe biblioteca standard Python + ezdxf + PyMuPDF (fără dependențe externe
precum shapely/pdfplumber), pentru integrare directă în SteelSync Toolkit și PyInstaller.
"""

from __future__ import annotations

import csv
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Optional

try:
    import ezdxf
    from ezdxf.enums import TextEntityAlignment
except ImportError:  # pragma: no cover
    ezdxf = None  # type: ignore[assignment]
    TextEntityAlignment = None  # type: ignore[assignment]

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf  # type: ignore[no-redef]
    except ImportError:
        pymupdf = None  # type: ignore[assignment]

try:
    from pdf_to_dwg import PdfToCadConverter
except ImportError:  # pragma: no cover
    PdfToCadConverter = None  # type: ignore[assignment]

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError:  # pragma: no cover
    openpyxl = None  # type: ignore[assignment]


DSTV_BLOCK_CODES = {
    "ST", "EN", "AK", "IK", "BO", "SI", "PU", "KO", "KA", "TO", "UE", "SC",
}
KNOWN_FACES = {"v", "h", "o", "u"}

CHESON_TYPE_LABELS = {
    "cheson_cu_talpa": "Tip 1: Cheson cu Talpă (Talpă + Capac + 2 Laterale)",
    "cheson_fara_talpa": "Tip 2: Cheson fără Talpă (2 Capace + 2 Laterale)",
    "cheson_dublu": "Tip 3: Cheson Dublu (Talpă Mijloc + 2 Capace + 4 Laterale)",
    "nominal_parametric": "Nominal HSQ din NC1",
}

_COORD_TOKEN_RE = re.compile(r"^([+-]?\d+(?:\.\d+)?)([a-zA-Z]*)$")
_PROFILE_3PART_RE = re.compile(
    r"^(HSQ|BOX|WQ)\s*(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)"
    r"-(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)"
    r"-(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)",
    re.IGNORECASE,
)
_PROFILE_2PART_RE = re.compile(
    r"^(HSQ|BOX|WQ)\s*(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)"
    r"-(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)(?:-[*\s]*)?$",
    re.IGNORECASE,
)
_PLATE_PROFILE_RE = re.compile(
    r"^(?:PL|PLT|PLATE|FL)\s*(\d+(?:[.,]\d+)?)\s*[*xX×]\s*(\d+(?:[.,]\d+)?)",
    re.IGNORECASE,
)
BOM_BANNER_KEYWORDS = (
    "STYKKLISTE",
    "STYCKLISTA",
    "STÜCKLISTE",
    "STUCKLISTE",
    "BILL OF MATERIALS",
    "PART LIST",
    "PARTS LIST",
)


# ─── Dataclasses ──────────────────────────────────────────────────────────────

@dataclass
class DSTVHeader:
    order_id: str = ""
    drawing_id: str = ""
    phase_id: str = ""
    piece_id: str = ""
    material: str = ""
    quantity: int = 1
    profile_str: str = ""
    profile_code: str = ""
    length: float = 0.0
    net_length: Optional[float] = None
    height: float = 0.0
    flange_width: float = 0.0
    flange_thickness: float = 0.0
    web_thickness: float = 0.0
    radius: float = 0.0
    weight_per_meter: float = 0.0
    surface_per_meter: float = 0.0
    web_cut_start: float = 0.0
    web_cut_end: float = 0.0
    flange_cut_start: float = 0.0
    flange_cut_end: float = 0.0
    text_lines: tuple[str, ...] = ()


@dataclass
class DSTVVertex:
    x: float
    y: float
    radius: float = 0.0
    x_flag: str = ""
    y_flag: str = ""
    bevel_top: float = 0.0
    bevel_top_depth: float = 0.0
    bevel_bot: float = 0.0
    bevel_bot_depth: float = 0.0
    is_corner_relief: bool = False


@dataclass
class DSTVContour:
    face: str
    vertices: list[DSTVVertex] = field(default_factory=list)
    block_type: str = "AK"


@dataclass
class DSTVHole:
    face: str
    x: float
    y: float
    diameter: float
    depth: float = 0.0
    hole_type: str = ""
    slot_length: float = 0.0
    slot_width: float = 0.0
    slot_angle: float = 0.0
    x_flag: str = ""
    y_flag: str = ""


@dataclass
class DSTVMarking:
    face: str
    x: float
    y: float
    angle: float = 0.0
    height: float = 10.0
    text: str = ""
    x_flag: str = ""
    y_flag: str = ""


@dataclass
class DSTVFile:
    header: DSTVHeader
    ak_contours: dict[str, list[DSTVContour]] = field(default_factory=dict)
    ik_contours: dict[str, list[DSTVContour]] = field(default_factory=dict)
    holes: dict[str, list[DSTVHole]] = field(default_factory=dict)
    markings: dict[str, list[DSTVMarking]] = field(default_factory=dict)
    source_path: Optional[Path] = None


@dataclass
class CompoundProfileSpec:
    profile_type: str
    h_web: float
    tw: float
    tf_bot: float
    b_bot: float
    tf_top: float
    b_top: float
    total_height: float
    total_width: float
    top_over_webs: bool
    bot_over_webs: bool


@dataclass
class ChesonConfig:
    cheson_type: str = "auto"               # auto | cheson_cu_talpa | cheson_fara_talpa | cheson_dublu | nominal_parametric
    flange_top_mode: str = "between_webs"   # auto | over_webs | between_webs
    flange_bot_mode: str = "between_webs"   # auto | over_webs | between_webs
    web_top_recess_mm: float | str = "auto_tw"
    web_bot_recess_mm: float | str = "auto_tw"
    web_skew_plane: str = "inner"           # outer | centerline | inner
    web_length_rounding_mm: float = 0.0
    length_offset_mm: float = 0.0
    exact_gabarit_lock: bool = True
    target_length_mm: Optional[float] = None
    override_web_height_mm: Optional[float] = None
    mirror_flange_o_y: bool = True
    mirror_flange_u_y: bool = True
    mirror_web_h_x: bool = False
    web_ordering: str = "upper_first_h_then_v"
    arc_segments_per_90deg: int = 16
    strip_cnc_corner_loops: bool = True
    corner_relief_radius_max_mm: float = 5.0
    strip_residual_fragments: bool = True
    deduplicate_identical: bool = True
    min_segment_length_mm: float = 0.05
    lowercase_piece_prefix: bool = True
    flange_letter: str = "C"
    web_letter: str = "L"

    @classmethod
    def from_preset(
        cls,
        preset: str,
        exact_gabarit_lock: bool = True,
        target_length_mm: Optional[float] = None,
        override_web_height_mm: Optional[float] = None,
        strip_residual_fragments: bool = True,
        deduplicate_identical: bool = True,
    ) -> "ChesonConfig":
        key = (preset or "auto").strip().lower()
        common = {
            "exact_gabarit_lock": exact_gabarit_lock,
            "target_length_mm": target_length_mm,
            "override_web_height_mm": override_web_height_mm,
            "strip_residual_fragments": strip_residual_fragments,
            "deduplicate_identical": deduplicate_identical,
        }
        if key in ("auto", "auto_detect"):
            return cls(cheson_type="auto", **common)
        if key in ("cheson_cu_talpa", "tip1"):
            return cls(cheson_type="cheson_cu_talpa", **common)
        if key in ("cheson_fara_talpa", "tip2"):
            return cls(cheson_type="cheson_fara_talpa", **common)
        if key in ("cheson_dublu", "tip3", "workshop_corner_weld"):
            return cls(cheson_type="cheson_dublu", **common)
        if key in ("nominal", "nominal_parametric"):
            return cls(
                cheson_type="nominal_parametric",
                flange_top_mode="auto",
                flange_bot_mode="auto",
                web_top_recess_mm=0.0,
                web_bot_recess_mm=0.0,
                web_skew_plane="outer",
                mirror_flange_o_y=False,
                mirror_flange_u_y=False,
                **common,
            )
        if key in ("box_both_over", "over_webs"):
            return cls(
                cheson_type="cheson_fara_talpa",
                flange_top_mode="over_webs",
                flange_bot_mode="over_webs",
                web_top_recess_mm=0.0,
                web_bot_recess_mm=0.0,
                web_skew_plane="outer",
                mirror_flange_o_y=True,
                mirror_flange_u_y=True,
                **common,
            )
        if key in ("box_both_between", "between_webs_no_recess"):
            return cls(
                cheson_type="cheson_fara_talpa",
                flange_top_mode="between_webs",
                flange_bot_mode="between_webs",
                web_top_recess_mm=0.0,
                web_bot_recess_mm=0.0,
                web_skew_plane="inner",
                mirror_flange_o_y=True,
                mirror_flange_u_y=True,
                **common,
            )
        return cls(cheson_type="auto", **common)


@dataclass
class ExtractedPlate:
    name: str
    folder: str
    role: str
    face: str
    thickness: float
    piece_id: str
    piece_clean: str
    material: str
    quantity: int
    exterior: list[tuple[float, float]]
    holes: list[DSTVHole] = field(default_factory=list)
    inner_contours: list[list[tuple[float, float]]] = field(default_factory=list)
    markings: list[DSTVMarking] = field(default_factory=list)
    source_bounds: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    removed_residuals: int = 0
    qty_per_cheson: int = 1
    role_code: str = ""
    merged_sources: list[str] = field(default_factory=list)

    @property
    def bounds(self) -> tuple[float, float, float, float]:
        return polygon_bounds(self.exterior)

    @property
    def area(self) -> float:
        return abs(polygon_area(self.exterior))


@dataclass
class VerificationReport:
    piece_id: str
    detected_type: str
    detected_type_label: str
    l_nc1_raw_mm: float
    l_gabarit_target_mm: float
    l_source_label: str
    l_assembled_mm: float
    h_target_mm: float
    h_assembled_mm: float
    w_target_mm: float
    w_assembled_mm: float
    max_gabarit_error_mm: float
    flange_web_contact_error_mm: float
    passed: bool


@dataclass
class DrawnHoleInfo:
    kind: str  # circle | slot | rect_slot
    cx: float
    cy: float
    width_mm: float
    length_mm: float
    angle_deg: float
    area_mm2: float


@dataclass
class DrawnDxfAudit:
    plate_name: str
    role: str
    face: str
    thickness_mm: float
    dxf_errors: int
    entity_layers: list[str]
    extra_layers_created: list[str]
    layer_zero_ok: bool
    outer_closed: bool
    origin_err_mm: float
    length_mm: float
    width_mm: float
    outer_area_mm2: float
    net_area_mm2: float
    vertex_count: int
    holes_drawn: list[DrawnHoleInfo]
    inners_count: int
    dim_err_mm: float
    hole_pos_err_mm: float
    hole_dim_err_mm: float
    holes_matched: int
    holes_expected: int
    passed: bool
    issues: list[str] = field(default_factory=list)


@dataclass
class Sim3DReport:
    l_3d_mm: float
    b_3d_mm: float
    h_3d_mm: float
    err_l_mm: float
    err_b_mm: float
    err_h_mm: float
    max_interference_mm: float
    interference_pairs: list[str]
    max_seam_gap_mm: float
    seam_notes: list[str]
    face_iou_pct: dict[str, float]
    overall_3d_similarity_pct: float
    holes_3d_expected: int
    holes_3d_matched: int
    max_hole_3d_err_mm: float
    hole_shape_mismatches: int
    mass_3d_kg: float
    mass_bom_kg: Optional[float]
    mass_diff_pct: Optional[float]
    passed: bool
    issues: list[str] = field(default_factory=list)


@dataclass
class PdfDrawingAnalysis:
    pdf_name: str
    text_source: str  # "vector" | "ocr_tesseract" | "ocr_winrt" | "none"
    dimensions: set[float] = field(default_factory=set)
    profile_callouts: list[str] = field(default_factory=list)
    hole_callouts: list[tuple[int, float]] = field(default_factory=list)
    raw_char_count: int = 0


@dataclass
class PdfOverlayReport:
    pdf_name: str
    top_scale: Optional[float]
    front_scale: Optional[float]
    top_overlay_pct: float
    front_overlay_pct: float
    overall_overlay_pct: float
    top_mean_err_mm: float
    front_mean_err_mm: float
    holes_pdf_expected: int
    holes_pdf_matched: int
    missing_structural_lines: int
    passed: bool
    issues: list[str] = field(default_factory=list)


@dataclass
class PostDrawVerificationReport:
    piece_id: str
    plates_audited: int
    all_dxf_clean: bool
    all_layer_zero: bool
    nc1_max_dim_err_mm: float
    nc1_max_hole_err_mm: float
    pdf_drawing_name: Optional[str]
    pdf_bom_length_err_mm: Optional[float]
    pdf_dims_confirmed: list[float]
    sim_3d: Sim3DReport
    plate_audits: list[DrawnDxfAudit]
    passed: bool
    issues: list[str] = field(default_factory=list)
    pdf_text_source: str = "none"
    pdf_profile_callouts: list[str] = field(default_factory=list)
    pdf_hole_callouts: list[tuple[int, float]] = field(default_factory=list)
    pdf_overlay: Optional[PdfOverlayReport] = None
    refinement_iterations: int = 0
    refinement_notes: list[str] = field(default_factory=list)
    snapped_noted_dims: list[str] = field(default_factory=list)


@dataclass
class BomRow:
    pos_no: str
    littra: str
    qty_per_assembly: int
    profile: str
    material: str
    length_mm: Optional[float]
    weight_unit_kg: Optional[float]
    weight_total_kg: Optional[float]


@dataclass
class BomTable:
    assembly_id: str
    assembly_quantity: int
    rows: list[BomRow]
    total_weight_kg: Optional[float]
    source_pdf: Path


@dataclass
class ManifestEntry:
    pos_no: str
    plate_name: str
    profile: str
    material: str
    thickness_mm: Optional[float]
    width_mm: Optional[float]
    length_mm: Optional[float]
    qty_per_assembly: int
    assembly_qty: int
    total_qty: int
    classification: str
    status: str
    source_file: str
    output_folder: str
    output_file: str
    notes: str = ""
    role_code: str = ""
    role_label: str = ""
    qty_per_cheson: int = 1
    qty_source: str = ""


# ─── Pure-Python 2D Polygon Engine ───────────────────────────────────────────

Ring = list[tuple[float, float]]


def polygon_area(ring: Iterable[tuple[float, float]]) -> float:
    pts = list(ring)
    n = len(pts)
    if n < 3:
        return 0.0
    s = 0.0
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return 0.5 * s


def polygon_bounds(ring: Iterable[tuple[float, float]]) -> tuple[float, float, float, float]:
    pts = list(ring)
    if not pts:
        return (0.0, 0.0, 0.0, 0.0)
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def _dist_point_to_segment(
    px: float, py: float, ax: float, ay: float, bx: float, by: float
) -> float:
    vx, vy = bx - ax, by - ay
    len_sq = vx * vx + vy * vy
    if len_sq < 1e-12:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * vx + (py - ay) * vy) / len_sq))
    proj_x = ax + t * vx
    proj_y = ay + t * vy
    return math.hypot(px - proj_x, py - proj_y)


def point_in_polygon(
    x: float, y: float, ring: Iterable[tuple[float, float]], tol: float = 0.05
) -> bool:
    pts = list(ring)
    n = len(pts)
    if n < 3:
        return False
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        if _dist_point_to_segment(x, y, x1, y1, x2, y2) <= tol:
            return True
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if ((yi > y) != (yj > y)) and (
            x < (xj - xi) * (y - yi) / ((yj - yi) if abs(yj - yi) > 1e-12 else 1e-12) + xi
        ):
            inside = not inside
        j = i
    return inside


def clean_ring(
    coords: Iterable[tuple[float, float]],
    min_seg_len: float = 0.05,
    collinear_tol: float = 1e-4,
    ccw: Optional[bool] = True,
) -> Ring:
    pts = [(round(float(x), 6), round(float(y), 6)) for x, y in coords]
    if len(pts) >= 2 and math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1]) < 1e-6:
        pts = pts[:-1]
    if len(pts) < 3:
        return pts

    filtered: Ring = [pts[0]]
    for pt in pts[1:]:
        if math.hypot(pt[0] - filtered[-1][0], pt[1] - filtered[-1][1]) >= min_seg_len:
            filtered.append(pt)
    if (
        len(filtered) >= 2
        and math.hypot(filtered[0][0] - filtered[-1][0], filtered[0][1] - filtered[-1][1]) < min_seg_len
    ):
        filtered.pop()

    if len(filtered) < 3:
        return filtered

    changed = True
    while changed and len(filtered) >= 3:
        changed = False
        new_pts: Ring = []
        n = len(filtered)
        for i in range(n):
            p_prev = filtered[(i - 1) % n]
            p_curr = filtered[i]
            p_next = filtered[(i + 1) % n]
            vx1, vy1 = p_curr[0] - p_prev[0], p_curr[1] - p_prev[1]
            vx2, vy2 = p_next[0] - p_curr[0], p_next[1] - p_curr[1]
            l1 = math.hypot(vx1, vy1)
            l2 = math.hypot(vx2, vy2)
            if l1 < min_seg_len or l2 < min_seg_len:
                changed = True
                continue
            cross = abs(vx1 * vy2 - vy1 * vx2) / (l1 * l2)
            dot = (vx1 * vx2 + vy1 * vy2) / (l1 * l2)
            if cross < collinear_tol and dot > 0.0:
                changed = True
                continue
            new_pts.append(p_curr)
        filtered = new_pts

    if ccw is not None and len(filtered) >= 3:
        area = polygon_area(filtered)
        if ccw and area < 0.0:
            filtered.reverse()
        elif (not ccw) and area > 0.0:
            filtered.reverse()

    return filtered


def _sutherland_hodgman_y(ring: Ring, threshold: float, keep_ge: bool) -> Ring:
    if len(ring) < 3:
        return []

    def _inside(p: tuple[float, float]) -> bool:
        return p[1] >= threshold if keep_ge else p[1] <= threshold

    def _intersect(p1: tuple[float, float], p2: tuple[float, float]) -> tuple[float, float]:
        dy = p2[1] - p1[1]
        if abs(dy) < 1e-12:
            return (p1[0], threshold)
        t = (threshold - p1[1]) / dy
        return (p1[0] + t * (p2[0] - p1[0]), threshold)

    out: Ring = []
    n = len(ring)
    for i in range(n):
        curr = ring[i]
        nxt = ring[(i + 1) % n]
        c_in = _inside(curr)
        n_in = _inside(nxt)
        if c_in and n_in:
            out.append(nxt)
        elif c_in and not n_in:
            out.append(_intersect(curr, nxt))
        elif (not c_in) and n_in:
            out.append(_intersect(curr, nxt))
            out.append(nxt)
    return out


def clip_ring_y_strip(
    ring: Ring,
    y_min: float,
    y_max: float,
    min_seg_len: float = 0.05,
) -> Ring:
    if len(ring) < 3 or y_max <= y_min:
        return []
    eps = 1e-5
    clipped = _sutherland_hodgman_y(ring, y_min + eps, keep_ge=True)
    clipped = _sutherland_hodgman_y(clipped, y_max - eps, keep_ge=False)
    if len(clipped) < 3:
        return []

    snapped: Ring = []
    for x, y in clipped:
        ny = y
        if abs(y - y_min) <= 2e-4:
            ny = y_min
        elif abs(y - y_max) <= 2e-4:
            ny = y_max
        snapped.append((x, ny))

    return clean_ring(snapped, min_seg_len=min_seg_len, ccw=True)


def mirror_ring(
    ring: Ring, cx: float, cy: float, mirror_x: bool = False, mirror_y: bool = False
) -> Ring:
    return [
        ((2.0 * cx - x) if mirror_x else x, (2.0 * cy - y) if mirror_y else y)
        for x, y in ring
    ]


def translate_ring(ring: Ring, dx: float, dy: float) -> Ring:
    return [(x + dx, y + dy) for x, y in ring]


def _cw_angle_from_back(dx_in: float, dy_in: float, dx_out: float, dy_out: float) -> float:
    ang_back = math.atan2(-dy_in, -dx_in)
    ang_out = math.atan2(dy_out, dx_out)
    cw = (ang_back - ang_out) % (2.0 * math.pi)
    return cw if cw > 1e-9 else 2.0 * math.pi


def _prune_thin_horizontal_slabs(
    ring: Ring,
    max_sliver_mm: float,
    min_seg_len: float = 0.05,
) -> tuple[list[Ring], int]:
    """Remove horizontal sliver strips (height <= max_sliver_mm) via vertical X-slab decomposition."""
    if len(ring) < 3 or max_sliver_mm <= 0.0:
        return [ring] if len(ring) >= 3 else [], 0

    minx, miny, maxx, maxy = polygon_bounds(ring)
    if min(maxx - minx, maxy - miny) <= max_sliver_mm:
        return [], 1

    raw_xs = sorted(p[0] for p in ring)
    xs: list[float] = [raw_xs[0]]
    for x in raw_xs[1:]:
        if x - xs[-1] > 1e-6:
            xs.append(x)
    if len(xs) < 2:
        return [ring], 0

    # (x_l, x_r, bot_yl, bot_yr, top_yl, top_yr)
    kept: list[tuple[float, float, float, float, float, float]] = []
    dropped_count = 0
    n = len(ring)

    for idx in range(len(xs) - 1):
        x_l, x_r = xs[idx], xs[idx + 1]
        dx = x_r - x_l
        if dx <= 1e-5:
            continue
        x_mid = 0.5 * (x_l + x_r)
        crossings: list[tuple[float, float, float, float]] = []
        for i in range(n):
            x1, y1 = ring[i]
            x2, y2 = ring[(i + 1) % n]
            lo_x, hi_x = (x1, x2) if x1 < x2 else (x2, x1)
            if lo_x < x_mid < hi_x:
                inv_dx = 1.0 / (x2 - x1)
                y_l = y1 + (x_l - x1) * (y2 - y1) * inv_dx
                y_m = y1 + (x_mid - x1) * (y2 - y1) * inv_dx
                y_r = y1 + (x_r - x1) * (y2 - y1) * inv_dx
                slope = (y2 - y1) * inv_dx
                crossings.append((y_m, y_l, y_r, slope))
        if len(crossings) < 2:
            continue
        crossings.sort(key=lambda c: c[0])
        for j in range(0, len(crossings) - 1, 2):
            _, bot_yl, bot_yr, bot_slope = crossings[j]
            _, top_yl, top_yr, top_slope = crossings[j + 1]
            h_l = top_yl - bot_yl
            h_r = top_yr - bot_yr
            is_sliver = (
                min(h_l, h_r) > 0.05
                and max(h_l, h_r) <= max_sliver_mm
                and max(abs(bot_slope), abs(top_slope)) < 0.25
            )
            if is_sliver:
                dropped_count += 1
            else:
                kept.append((x_l, x_r, bot_yl, bot_yr, top_yl, top_yr))

    if dropped_count == 0:
        return [ring], 0
    if not kept:
        return [], dropped_count

    def _vkey(pt: tuple[float, float]) -> tuple[float, float]:
        return (round(pt[0], 4), round(pt[1], 4))

    edges: list[tuple[tuple[float, float], tuple[float, float]]] = []
    slab_xs_raw = sorted({t[0] for t in kept} | {t[1] for t in kept})
    slab_xs: list[float] = []
    for x in slab_xs_raw:
        if not slab_xs or x - slab_xs[-1] > 1e-6:
            slab_xs.append(x)

    for x_l, x_r, bot_yl, bot_yr, top_yl, top_yr in kept:
        edges.append(((x_l, bot_yl), (x_r, bot_yr)))
        edges.append(((x_r, top_yr), (x_l, top_yl)))

    for x_k in slab_xs:
        up_iv = [
            (byr, tyr)
            for xl, xr, _, byr, _, tyr in kept
            if abs(xr - x_k) <= 1e-6 and tyr - byr > 1e-6
        ]
        dn_iv = [
            (byl, tyl)
            for xl, xr, byl, _, tyl, _ in kept
            if abs(xl - x_k) <= 1e-6 and tyl - byl > 1e-6
        ]
        if not up_iv and not dn_iv:
            continue
        ys_raw = sorted(
            [v for a, b in up_iv for v in (a, b)]
            + [v for a, b in dn_iv for v in (a, b)]
        )
        ys_k: list[float] = [ys_raw[0]]
        for yv in ys_raw[1:]:
            if yv - ys_k[-1] > 1e-6:
                ys_k.append(yv)

        up_segs: list[list[float]] = []
        dn_segs: list[list[float]] = []
        for idx_y in range(len(ys_k) - 1):
            ya, yb = ys_k[idx_y], ys_k[idx_y + 1]
            if yb - ya <= 1e-6:
                continue
            ym = 0.5 * (ya + yb)
            net = sum(1 for a, b in up_iv if a < ym < b) - sum(1 for a, b in dn_iv if a < ym < b)
            if net > 0:
                if up_segs and abs(up_segs[-1][1] - ya) <= 1e-6:
                    up_segs[-1][1] = yb
                else:
                    up_segs.append([ya, yb])
            elif net < 0:
                if dn_segs and abs(dn_segs[-1][1] - ya) <= 1e-6:
                    dn_segs[-1][1] = yb
                else:
                    dn_segs.append([ya, yb])

        for ya, yb in up_segs:
            edges.append(((x_k, ya), (x_k, yb)))
        for ya, yb in dn_segs:
            edges.append(((x_k, yb), (x_k, ya)))

    if len(edges) < 3:
        return [], dropped_count

    adj: dict[tuple[float, float], list[int]] = {}
    for e_idx, (p_start, _) in enumerate(edges):
        adj.setdefault(_vkey(p_start), []).append(e_idx)

    used = [False] * len(edges)
    loops: list[Ring] = []

    for start_idx in range(len(edges)):
        if used[start_idx]:
            continue
        loop_pts: Ring = []
        curr_idx = start_idx
        start_key = _vkey(edges[start_idx][0])
        steps = 0
        max_steps = len(edges) + 2
        closed = False

        while steps < max_steps:
            steps += 1
            used[curr_idx] = True
            p_a, p_b = edges[curr_idx]
            loop_pts.append(p_a)
            next_key = _vkey(p_b)
            if next_key == start_key:
                closed = True
                break
            candidates = [i for i in adj.get(next_key, []) if not used[i]]
            if not candidates:
                break
            if len(candidates) == 1:
                curr_idx = candidates[0]
            else:
                dx_in = p_b[0] - p_a[0]
                dy_in = p_b[1] - p_a[1]
                curr_idx = min(
                    candidates,
                    key=lambda ci: _cw_angle_from_back(
                        dx_in,
                        dy_in,
                        edges[ci][1][0] - edges[ci][0][0],
                        edges[ci][1][1] - edges[ci][0][1],
                    ),
                )

        if closed and len(loop_pts) >= 3 and polygon_area(loop_pts) > 0.1:
            cleaned = clean_ring(loop_pts, min_seg_len=min_seg_len, ccw=True)
            if len(cleaned) >= 3 and abs(polygon_area(cleaned)) > 0.1:
                b = polygon_bounds(cleaned)
                if min(b[2] - b[0], b[3] - b[1]) <= max_sliver_mm:
                    dropped_count += 1
                else:
                    loops.append(cleaned)

    return loops, dropped_count


def remove_residual_fragments(
    ring: Ring,
    plate_thickness_mm: float,
    min_seg_len: float = 0.05,
) -> tuple[list[Ring], int]:
    """Strip residual ears/tabs whose thickness is <= plate_thickness_mm (both axes)."""
    if len(ring) < 3 or plate_thickness_mm <= 0.0:
        return [ring] if len(ring) >= 3 else [], 0

    max_sliver = float(plate_thickness_mm) + 0.01
    rings_x, rem_x = _prune_thin_horizontal_slabs(ring, max_sliver, min_seg_len=min_seg_len)

    rings_xy: list[Ring] = []
    rem_y = 0
    for r in rings_x:
        # Reflect (x, y) -> (y, x) and reverse order so CCW orientation is preserved
        swapped = [(y, x) for x, y in reversed(r)]
        sub_y, cnt_y = _prune_thin_horizontal_slabs(swapped, max_sliver, min_seg_len=min_seg_len)
        rem_y += cnt_y
        for sr in sub_y:
            restored = [(y, x) for x, y in reversed(sr)]
            cleaned = clean_ring(restored, min_seg_len=min_seg_len, ccw=True)
            if len(cleaned) >= 3 and abs(polygon_area(cleaned)) > 0.1:
                rings_xy.append(cleaned)

    if rem_x == 0 and rem_y == 0:
        return [ring], 0
    return rings_xy, rem_x + rem_y


def clean_piece_name(piece_id: str, lowercase: bool = True) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]", "", piece_id or "")
    if not cleaned:
        cleaned = "part"
    return cleaned.lower() if lowercase else cleaned.upper()


# ─── DSTV (.nc1) Parser ──────────────────────────────────────────────────────

def _parse_coord_token(token: str) -> tuple[float, str]:
    m = _COORD_TOKEN_RE.match(token.strip())
    if not m:
        raise ValueError(f"Invalid DSTV numeric/flag token: {token!r}")
    return float(m.group(1)), m.group(2).lower()


def _parse_float(raw: str, default: float = 0.0) -> float:
    s = (raw or "").strip()
    if not s:
        return default
    return float(s.split(",")[0].strip())


def _parse_int(raw: str, default: int = 1) -> int:
    s = (raw or "").strip()
    if not s:
        return default
    return int(float(s))


def _is_block_marker(line: str) -> Optional[str]:
    stripped = line.strip()
    if len(stripped) == 2 and stripped.isalpha():
        code = stripped.upper()
        if code in DSTV_BLOCK_CODES:
            return code
    return None


def _parse_header_lines(lines: list[str]) -> DSTVHeader:
    def get(idx: int) -> str:
        return lines[idx].strip() if idx < len(lines) else ""

    len_raw = get(8)
    length = 0.0
    net_length: Optional[float] = None
    if len_raw:
        parts = [p.strip() for p in len_raw.split(",") if p.strip()]
        if parts:
            length = float(parts[0])
        if len(parts) > 1:
            net_length = float(parts[1])

    text_lines = tuple(get(i) for i in range(20, min(len(lines), 24)))
    return DSTVHeader(
        order_id=get(0),
        drawing_id=get(1),
        phase_id=get(2),
        piece_id=get(3),
        material=get(4),
        quantity=_parse_int(get(5), 1),
        profile_str=get(6),
        profile_code=get(7),
        length=length,
        net_length=net_length,
        height=_parse_float(get(9)),
        flange_width=_parse_float(get(10)),
        flange_thickness=_parse_float(get(11)),
        web_thickness=_parse_float(get(12)),
        radius=_parse_float(get(13)),
        weight_per_meter=_parse_float(get(14)),
        surface_per_meter=_parse_float(get(15)),
        web_cut_start=_parse_float(get(16)),
        web_cut_end=_parse_float(get(17)),
        flange_cut_start=_parse_float(get(18)),
        flange_cut_end=_parse_float(get(19)),
        text_lines=text_lines,
    )


def parse_dstv_text(content: str, source_path: Optional[Path] = None) -> DSTVFile:
    raw_lines = content.splitlines()
    idx = 0
    n = len(raw_lines)

    while idx < n:
        stripped = raw_lines[idx].strip()
        idx += 1
        if stripped.upper() == "ST":
            break
    else:
        raise ValueError("Missing 'ST' start marker in DSTV file")

    header_lines: list[str] = []
    while idx < n and len(header_lines) < 24:
        line = raw_lines[idx]
        if line.lstrip().startswith("**"):
            idx += 1
            continue
        if _is_block_marker(line) is not None:
            break
        header_lines.append(line)
        idx += 1

    header = _parse_header_lines(header_lines)

    ak_contours: dict[str, list[DSTVContour]] = {}
    ik_contours: dict[str, list[DSTVContour]] = {}
    holes: dict[str, list[DSTVHole]] = {}
    markings: dict[str, list[DSTVMarking]] = {}

    current_block: Optional[str] = None
    current_contour: Optional[DSTVContour] = None
    last_face = "v"

    def _flush_contour() -> None:
        nonlocal current_contour
        if current_contour is not None and current_contour.vertices:
            target = ak_contours if current_contour.block_type == "AK" else ik_contours
            target.setdefault(current_contour.face, []).append(current_contour)
        current_contour = None

    while idx < n:
        line = raw_lines[idx]
        idx += 1
        stripped = line.strip()
        if not stripped or stripped.startswith("**"):
            continue

        marker = _is_block_marker(line)
        if marker is not None:
            _flush_contour()
            if marker == "EN":
                break
            current_block = marker
            continue

        if current_block in ("AK", "IK"):
            tokens = stripped.split()
            if not tokens:
                continue
            if tokens[0].lower() in KNOWN_FACES:
                face = tokens[0].lower()
                last_face = face
                tokens = tokens[1:]
                if current_contour is None:
                    current_contour = DSTVContour(face=face, block_type=current_block)
                elif current_contour.face != face:
                    _flush_contour()
                    current_contour = DSTVContour(face=face, block_type=current_block)
            elif current_contour is None:
                current_contour = DSTVContour(face=last_face, block_type=current_block)

            if len(tokens) < 2:
                continue
            x_val, x_flag = _parse_coord_token(tokens[0])
            y_val, y_flag = _parse_coord_token(tokens[1])
            rad = _parse_coord_token(tokens[2])[0] if len(tokens) > 2 else 0.0
            b_top = _parse_coord_token(tokens[3])[0] if len(tokens) > 3 else 0.0
            b_top_d = _parse_coord_token(tokens[4])[0] if len(tokens) > 4 else 0.0
            b_bot = _parse_coord_token(tokens[5])[0] if len(tokens) > 5 else 0.0
            b_bot_d = _parse_coord_token(tokens[6])[0] if len(tokens) > 6 else 0.0

            current_contour.vertices.append(
                DSTVVertex(
                    x=x_val,
                    y=y_val,
                    radius=rad,
                    x_flag=x_flag,
                    y_flag=y_flag,
                    bevel_top=b_top,
                    bevel_top_depth=b_top_d,
                    bevel_bot=b_bot,
                    bevel_bot_depth=b_bot_d,
                    is_corner_relief=(x_flag == "w" or y_flag == "w"),
                )
            )

        elif current_block == "BO":
            tokens = stripped.split()
            if not tokens:
                continue
            if tokens[0].lower() in KNOWN_FACES:
                face = tokens[0].lower()
                last_face = face
                tokens = tokens[1:]
            else:
                face = last_face
            if len(tokens) < 3:
                continue
            x_val, x_flag = _parse_coord_token(tokens[0])
            y_val, y_flag = _parse_coord_token(tokens[1])
            diam, _ = _parse_coord_token(tokens[2])
            depth = 0.0
            hole_type = ""
            slot_len = 0.0
            slot_w = 0.0
            slot_ang = 0.0
            rest = tokens[3:]
            if rest:
                if rest[0].lower() in ("l", "g", "m", "k"):
                    hole_type = rest[0].lower()
                    rest = rest[1:]
                else:
                    d_val, d_flag = _parse_coord_token(rest[0])
                    depth = d_val
                    if d_flag:
                        hole_type = d_flag
                    rest = rest[1:]
                    if rest and rest[0].lower() in ("l", "g", "m", "k"):
                        hole_type = rest[0].lower()
                        rest = rest[1:]
            if hole_type == "l" and len(rest) >= 1:
                slot_len = _parse_coord_token(rest[0])[0]
                if len(rest) >= 2:
                    slot_w = _parse_coord_token(rest[1])[0]
                if len(rest) >= 3:
                    slot_ang = _parse_coord_token(rest[2])[0]

            holes.setdefault(face, []).append(
                DSTVHole(
                    face=face,
                    x=x_val,
                    y=y_val,
                    diameter=diam,
                    depth=depth,
                    hole_type=hole_type,
                    slot_length=slot_len,
                    slot_width=slot_w,
                    slot_angle=slot_ang,
                    x_flag=x_flag,
                    y_flag=y_flag,
                )
            )

        elif current_block == "SI":
            m = re.match(
                r"^\s*(?:([vhrouVHROU])\s+)?(\S+)\s+(\S+)\s+(\S+)\s+(\S+)(?:\s+(.*))?$",
                line,
            )
            if not m:
                continue
            face = (m.group(1) or last_face).lower()
            last_face = face
            x_val, x_flag = _parse_coord_token(m.group(2))
            y_val, y_flag = _parse_coord_token(m.group(3))
            ang = _parse_coord_token(m.group(4))[0]
            h_txt = _parse_coord_token(m.group(5))[0]
            txt = (m.group(6) or "").strip()
            markings.setdefault(face, []).append(
                DSTVMarking(
                    face=face,
                    x=x_val,
                    y=y_val,
                    angle=ang,
                    height=h_txt,
                    text=txt,
                    x_flag=x_flag,
                    y_flag=y_flag,
                )
            )

    _flush_contour()
    return DSTVFile(
        header=header,
        ak_contours=ak_contours,
        ik_contours=ik_contours,
        holes=holes,
        markings=markings,
        source_path=source_path,
    )


def parse_dstv(nc1_path: str | Path) -> DSTVFile:
    path = Path(nc1_path)
    content = path.read_text(encoding="utf-8", errors="replace")
    return parse_dstv_text(content, source_path=path)


def _strip_cnc_corner_relief_loops(
    vertices: list[DSTVVertex],
    max_relief_radius: float = 5.0,
) -> list[DSTVVertex]:
    if len(vertices) < 4:
        return list(vertices)

    work = list(vertices)
    if (
        math.hypot(work[0].x - work[-1].x, work[0].y - work[-1].y) < 1e-6
        and abs(work[-1].radius) < 1e-6
    ):
        work = work[:-1]

    out: list[DSTVVertex] = []
    i = 0
    n = len(work)
    while i < n:
        if i + 3 < n:
            v0, v1, v2, v3 = work[i], work[i + 1], work[i + 2], work[i + 3]
            r_abs = abs(v1.radius)
            if (
                (v1.y_flag == "w" or v1.x_flag == "w")
                and 0.0 < r_abs <= max_relief_radius
                and abs(v0.radius) > 0.0
                and abs(v2.radius) > 0.0
                and abs(v3.radius) < 1e-6
            ):
                tol = r_abs * math.sqrt(2.0) + 0.2
                if (
                    math.hypot(v0.x - v1.x, v0.y - v1.y) <= tol
                    and math.hypot(v2.x - v1.x, v2.y - v1.y) <= tol
                    and math.hypot(v3.x - v1.x, v3.y - v1.y) <= tol
                ):
                    out.append(DSTVVertex(x=v1.x, y=v1.y, radius=0.0, x_flag=v1.x_flag, y_flag=v1.y_flag))
                    i += 4
                    continue

        if i + 2 < n:
            v0, v1, v2 = work[i], work[i + 1], work[i + 2]
            r_abs = abs(v1.radius)
            if (
                (v1.y_flag == "w" or v1.x_flag == "w")
                and 0.0 < r_abs <= max_relief_radius
                and abs(v0.radius) > 0.0
                and abs(v2.radius) < 1e-6
            ):
                tol = r_abs * math.sqrt(2.0) + 0.2
                if (
                    math.hypot(v0.x - v1.x, v0.y - v1.y) <= tol
                    and math.hypot(v2.x - v1.x, v2.y - v1.y) <= tol
                ):
                    out.append(DSTVVertex(x=v1.x, y=v1.y, radius=0.0, x_flag=v1.x_flag, y_flag=v1.y_flag))
                    i += 3
                    continue

        out.append(work[i])
        i += 1
    return out


def _discretize_dstv_arc(
    p1: tuple[float, float],
    p2: tuple[float, float],
    radius: float,
    arc_segments_per_90deg: int = 16,
) -> list[tuple[float, float]]:
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    chord = math.hypot(dx, dy)
    r_abs = abs(radius)
    if chord < 1e-6 or r_abs < 1e-6 or chord > 2.0 * r_abs + 0.05:
        return [p2]

    half_c = chord * 0.5
    h = math.sqrt(max(0.0, r_abs * r_abs - half_c * half_c))
    mx = 0.5 * (p1[0] + p2[0])
    my = 0.5 * (p1[1] + p2[1])
    nx = -dy / chord
    ny = dx / chord

    if radius > 0:
        cx, cy = mx + h * nx, my + h * ny
        a1 = math.atan2(p1[1] - cy, p1[0] - cx)
        a2 = math.atan2(p2[1] - cy, p2[0] - cx)
        dtheta = (a2 - a1) % (2.0 * math.pi)
        if dtheta < 1e-9:
            dtheta = math.pi
    else:
        cx, cy = mx - h * nx, my - h * ny
        a1 = math.atan2(p1[1] - cy, p1[0] - cx)
        a2 = math.atan2(p2[1] - cy, p2[0] - cx)
        dtheta = -((a1 - a2) % (2.0 * math.pi))
        if abs(dtheta) < 1e-9:
            dtheta = -math.pi

    n_segs = max(2, int(math.ceil(abs(dtheta) / (0.5 * math.pi) * arc_segments_per_90deg)))
    pts: list[tuple[float, float]] = []
    for k in range(1, n_segs):
        ang = a1 + dtheta * (k / n_segs)
        pts.append((cx + r_abs * math.cos(ang), cy + r_abs * math.sin(ang)))
    pts.append(p2)
    return pts


def contour_to_ring(
    contour: DSTVContour,
    cfg: Optional[ChesonConfig] = None,
) -> Ring:
    cfg = cfg or ChesonConfig()
    verts = contour.vertices
    if not verts:
        return []

    if cfg.strip_cnc_corner_loops:
        verts = _strip_cnc_corner_relief_loops(
            verts, max_relief_radius=cfg.corner_relief_radius_max_mm
        )
    else:
        if (
            len(verts) >= 2
            and math.hypot(verts[0].x - verts[-1].x, verts[0].y - verts[-1].y) < 1e-6
            and abs(verts[-1].radius) < 1e-6
        ):
            verts = verts[:-1]

    if len(verts) < 3:
        return []

    ring: Ring = [(verts[0].x, verts[0].y)]
    n = len(verts)
    for i in range(n):
        v_curr = verts[i]
        v_next = verts[(i + 1) % n]
        p1 = (v_curr.x, v_curr.y)
        p2 = (v_next.x, v_next.y)
        if math.hypot(p2[0] - p1[0], p2[1] - p1[1]) < 1e-7:
            continue
        if abs(v_curr.radius) > 1e-6:
            ring.extend(
                _discretize_dstv_arc(p1, p2, v_curr.radius, cfg.arc_segments_per_90deg)
            )
        else:
            ring.append(p2)

    return clean_ring(ring, min_seg_len=cfg.min_segment_length_mm, ccw=True)


def parse_compound_profile(profile_str: str, header: DSTVHeader) -> CompoundProfileSpec:
    s = (profile_str or "").strip()
    m3 = _PROFILE_3PART_RE.match(s)
    if m3:
        ptype = m3.group(1).upper()
        h_web = float(m3.group(2))
        tw = float(m3.group(3))
        tf_top = float(m3.group(4))
        b_top = float(m3.group(5))
        tf_bot = float(m3.group(6))
        b_bot = float(m3.group(7))
    else:
        m2 = _PROFILE_2PART_RE.match(s)
        if m2:
            ptype = m2.group(1).upper()
            h_web = float(m2.group(2))
            tw = float(m2.group(3))
            tf_bot = float(m2.group(4))
            b_bot = float(m2.group(5))
            tf_top = header.flange_thickness if header.flange_thickness > 0 else tf_bot
            b_top = header.flange_width if header.flange_width > 0 else b_bot
        else:
            ptype = "UNKNOWN"
            tw = header.web_thickness
            tf_bot = header.flange_thickness
            tf_top = header.flange_thickness
            b_top = header.flange_width
            b_bot = header.flange_width
            h_web = max(0.0, header.height - tf_top)

    B = header.flange_width if header.flange_width > 0 else max(b_top, b_bot)
    H = header.height if header.height > 0 else (h_web + tf_top)

    top_over_webs = abs(b_top - B) <= 0.5
    if abs(b_bot - (B - 2.0 * tw)) <= 0.5:
        bot_over_webs = False
    else:
        bot_over_webs = abs(b_bot - B) <= 0.5

    return CompoundProfileSpec(
        profile_type=ptype,
        h_web=h_web,
        tw=tw,
        tf_bot=tf_bot,
        b_bot=b_bot,
        tf_top=tf_top,
        b_top=b_top,
        total_height=H,
        total_width=B,
        top_over_webs=top_over_webs,
        bot_over_webs=bot_over_webs,
    )


def resolve_recess_mm(value: float | str, tw: float) -> float:
    if isinstance(value, str):
        s = value.strip().lower()
        if s == "auto_tw":
            return float(tw)
        return float(s)
    return float(value)


def _resolve_flange_over_webs(mode: str, spec_default: bool) -> bool:
    m = (mode or "auto").strip().lower()
    if m == "over_webs":
        return True
    if m == "between_webs":
        return False
    return spec_default


def _has_vertical_web_gap(contours: list[DSTVContour], cfg: ChesonConfig) -> bool:
    if len(contours) < 2:
        return False
    ranges = []
    for c in contours:
        r = contour_to_ring(c, cfg)
        if len(r) >= 3:
            _, miny, _, maxy = polygon_bounds(r)
            ranges.append((miny, maxy))
    if len(ranges) < 2:
        return False
    ranges.sort(key=lambda t: t[0])
    for i in range(len(ranges) - 1):
        if ranges[i + 1][0] - ranges[i][1] > 1.0:
            return True
    return False


def _face_y_span(dstv: DSTVFile, face: str, cfg: ChesonConfig) -> Optional[float]:
    contours = dstv.ak_contours.get(face, [])
    if not contours:
        return None
    ys: list[float] = []
    for c in contours:
        r = contour_to_ring(c, cfg)
        for _, y in r:
            ys.append(y)
    return (max(ys) - min(ys)) if ys else None


def _align_spec_flanges_to_faces(
    dstv: DSTVFile,
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
) -> CompoundProfileSpec:
    """Align top/bottom flange specs with actual .nc1 AK face widths (dy_o vs dy_u)."""
    dy_o = _face_y_span(dstv, "o", cfg)
    dy_u = _face_y_span(dstv, "u", cfg)
    if dy_o is None or dy_u is None:
        return spec
    thresh = 2.0 * spec.tw + 1.0
    if abs(dy_o - dy_u) <= thresh:
        return spec
    if abs(spec.b_top - spec.b_bot) <= 1.0:
        return spec
    o_is_wider = dy_o > dy_u + thresh
    spec_top_is_wider = spec.b_top > spec.b_bot + 1.0
    if o_is_wider != spec_top_is_wider:
        return CompoundProfileSpec(
            profile_type=spec.profile_type,
            h_web=spec.h_web,
            tw=spec.tw,
            tf_bot=spec.tf_top,
            b_bot=spec.b_top,
            tf_top=spec.tf_bot,
            b_top=spec.b_bot,
            total_height=spec.total_height,
            total_width=spec.total_width,
            top_over_webs=spec.bot_over_webs,
            bot_over_webs=spec.top_over_webs,
        )
    return spec


def detect_cheson_type(
    dstv: DSTVFile,
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
) -> tuple[str, str]:
    """Classify box beam into (cheson_type, talpa_side: 'none'|'bottom'|'top'|'middle')."""
    explicit = (cfg.cheson_type or "auto").strip().lower()
    tw = spec.tw

    dy_o = _face_y_span(dstv, "o", cfg)
    dy_u = _face_y_span(dstv, "u", cfg)
    if dy_o is not None and dy_u is not None and abs(dy_u - dy_o) > 1.0:
        bot_wider = dy_u > dy_o + 2.0 * tw + 1.0
        top_wider = dy_o > dy_u + 2.0 * tw + 1.0
    else:
        bot_wider = spec.b_bot > spec.b_top + 2.0 * tw + 1.0
        top_wider = spec.b_top > spec.b_bot + 2.0 * tw + 1.0
    talpa_side = "top" if top_wider and not bot_wider else ("bottom" if bot_wider else "none")

    if explicit != "auto":
        if explicit == "cheson_cu_talpa" and talpa_side == "none":
            talpa_side = "bottom"
        return explicit, talpa_side

    # Auto-detect from geometry
    if _has_vertical_web_gap(dstv.ak_contours.get("v", []), cfg) or _has_vertical_web_gap(
        dstv.ak_contours.get("h", []), cfg
    ):
        return "cheson_dublu", "middle"

    if bot_wider or top_wider:
        return "cheson_cu_talpa", talpa_side

    return "cheson_fara_talpa", "none"


def _apply_length_offset(ring: Ring, length_offset_mm: float) -> Ring:
    if not ring or abs(length_offset_mm) <= 1e-6:
        return ring
    _, _, maxx, _ = polygon_bounds(ring)
    return [(x + length_offset_mm if abs(x - maxx) < 0.5 else x, y) for x, y in ring]


def _compute_raw_x_envelope(
    rings_by_face: dict[str, list[Ring]]
) -> tuple[float, float, float]:
    xs: list[float] = []
    for rlist in rings_by_face.values():
        for r in rlist:
            for x, _ in r:
                xs.append(x)
    if not xs:
        return (0.0, 0.0, 0.0)
    x0, x1 = min(xs), max(xs)
    return (x0, x1, x1 - x0)


def _detect_face_end_skew(ring: Ring) -> tuple[bool, bool]:
    """Return (left_skewed, right_skewed) for a contour ring."""
    if len(ring) < 3:
        return False, False
    minx, _, maxx, _ = polygon_bounds(ring)
    span_x = max(1.0, maxx - minx)
    left_skew = False
    right_skew = False
    n = len(ring)
    for i in range(n):
        p1, p2 = ring[i], ring[(i + 1) % n]
        dx = abs(p2[0] - p1[0])
        dy = abs(p2[1] - p1[1])
        if dy > 20.0 and dx > 0.5:
            x_mid = 0.5 * (p1[0] + p2[0])
            if x_mid < minx + 0.25 * span_x:
                left_skew = True
            elif x_mid > maxx - 0.25 * span_x:
                right_skew = True
    return left_skew, right_skew


def _resolve_gabarit_length(
    l_nc1: float,
    cfg: ChesonConfig,
) -> tuple[float, str]:
    if cfg.target_length_mm is not None and cfg.target_length_mm > 0.0:
        return float(cfg.target_length_mm), "STYKKLISTE / Manual"
    if cfg.exact_gabarit_lock and abs(l_nc1 - round(l_nc1)) <= 1.0:
        return float(round(l_nc1)), "NC1_Snap_Int"
    return l_nc1, "NC1_Raw"


def _apply_gabarit_x_lock_to_rings(
    rings_by_face: dict[str, list[Ring]],
    x_min_env: float,
    l_nc1: float,
    l_gabarit: float,
    max_delta_mm: float = 20.0,
) -> dict[str, list[Ring]]:
    """Shift right-half vertices by delta_L and snap non-skewed left boundary jitter (<=0.25mm) to x_min_env."""
    if l_nc1 <= 1e-3:
        return rings_by_face
    delta_l = l_gabarit - l_nc1
    apply_delta = abs(delta_l) <= max_delta_mm
    mid_x = x_min_env + 0.5 * l_nc1
    x_max_target = x_min_env + (l_gabarit if apply_delta else l_nc1)

    # Determine if left/right ends have intentional skew on flanges
    flange_left_skew = False
    flange_right_skew = False
    for f in ("o", "u"):
        for r in rings_by_face.get(f, []):
            ls, rs = _detect_face_end_skew(r)
            flange_left_skew = flange_left_skew or ls
            flange_right_skew = flange_right_skew or rs

    out: dict[str, list[Ring]] = {}
    for face, rlist in rings_by_face.items():
        new_list: list[Ring] = []
        for r in rlist:
            ls, rs = _detect_face_end_skew(r)
            new_r: Ring = []
            for x, y in r:
                nx = x
                if apply_delta and x > mid_x:
                    nx = x + delta_l
                elif (not flange_left_skew) and (not ls) and abs(x - x_min_env) <= 0.25:
                    nx = x_min_env
                if (not flange_right_skew) and (not rs) and abs(nx - x_max_target) <= 0.25:
                    nx = x_max_target
                new_r.append((nx, y))
            new_list.append(clean_ring(new_r, min_seg_len=0.05, ccw=True))
        out[face] = new_list
    return out


def _estimate_end_skew_angles(
    dstv: DSTVFile,
    rings_by_face: dict[str, list[Ring]],
) -> tuple[float, float]:
    h = dstv.header
    start_tan = abs(math.tan(math.radians(h.flange_cut_start))) if abs(h.flange_cut_start) > 1e-4 else 0.0
    end_tan = abs(math.tan(math.radians(h.flange_cut_end))) if abs(h.flange_cut_end) > 1e-4 else 0.0

    if start_tan == 0.0 or end_tan == 0.0:
        for face in ("o", "u"):
            for ring in rings_by_face.get(face, []):
                if len(ring) < 3:
                    continue
                minx, miny, maxx, maxy = polygon_bounds(ring)
                n = len(ring)
                for i in range(n):
                    p1, p2 = ring[i], ring[(i + 1) % n]
                    dy = abs(p2[1] - p1[1])
                    dx = abs(p2[0] - p1[0])
                    if (
                        dy > 20.0
                        and dx > 0.5
                        and min(p1[1], p2[1]) <= miny + 1.0
                        and max(p1[1], p2[1]) >= maxy - 1.0
                    ):
                        if start_tan == 0.0 and min(p1[0], p2[0]) <= minx + 1.0:
                            start_tan = dx / dy
                        elif end_tan == 0.0 and max(p1[0], p2[0]) >= maxx - 1.0:
                            end_tan = dx / dy
    return start_tan, end_tan


def _flange_skew_x_contacts(
    flange_rings: list[Ring],
    y_low: float,
    y_high: float,
) -> tuple[Optional[tuple[float, float]], Optional[tuple[float, float]]]:
    """Return exact (left_x_min, left_x_max) and (right_x_min, right_x_max) on clipped flange."""
    if not flange_rings:
        return None, None
    # Use the first clipped flange contour
    r = flange_rings[0]
    minx, miny, maxx, maxy = polygon_bounds(r)
    mid_x = 0.5 * (minx + maxx)
    left_xs = [x for x, y in r if x <= mid_x and (abs(y - y_low) < 0.5 or abs(y - y_high) < 0.5)]
    right_xs = [x for x, y in r if x > mid_x and (abs(y - y_low) < 0.5 or abs(y - y_high) < 0.5)]
    left_pair = (min(left_xs), max(left_xs)) if len(left_xs) >= 2 else None
    right_pair = (min(right_xs), max(right_xs)) if len(right_xs) >= 2 else None
    return left_pair, right_pair


def _adjust_web_skew_and_rounding(
    ring: Ring,
    face: str,
    tw: float,
    start_tan: float,
    end_tan: float,
    skew_plane: str,
    rounding_mm: float,
    sync_left_x: Optional[float] = None,
    sync_right_x: Optional[float] = None,
    is_long_left: Optional[bool] = None,
    is_long_right: Optional[bool] = None,
) -> Ring:
    if len(ring) < 3:
        return ring

    plane = (skew_plane or "outer").strip().lower()
    k = 0.5 if plane == "centerline" else (1.0 if plane == "inner" else 0.0)

    minx, miny, maxx, maxy = polygon_bounds(ring)
    if sync_left_x is not None and start_tan > 1e-6:
        dx_left = sync_left_x - minx
    elif start_tan > 1e-6 and k > 0.0 and is_long_left is False:
        dx_left = -k * tw * start_tan
    else:
        dx_left = 0.0

    if sync_right_x is not None and end_tan > 1e-6:
        dx_right = sync_right_x - maxx
    elif end_tan > 1e-6 and k > 0.0 and is_long_right is False:
        dx_right = +k * tw * end_tan
    else:
        dx_right = 0.0

    if abs(dx_left) > 1e-7 or abs(dx_right) > 1e-7:
        span = maxx - minx
        left_cut = minx + 0.35 * span
        right_cut = maxx - 0.35 * span
        top_pts_l = [x for x, y in ring if abs(y - maxy) <= 1.0 and x <= left_cut]
        bot_pts_l = [x for x, y in ring if abs(y - miny) <= 1.0 and x <= left_cut]
        left_staggered = bool(top_pts_l and bot_pts_l and abs(min(top_pts_l) - min(bot_pts_l)) > 1.0)
        top_pts_r = [x for x, y in ring if abs(y - maxy) <= 1.0 and x >= right_cut]
        bot_pts_r = [x for x, y in ring if abs(y - miny) <= 1.0 and x >= right_cut]
        right_staggered = bool(top_pts_r and bot_pts_r and abs(max(top_pts_r) - max(bot_pts_r)) > 1.0)

        new_pts: Ring = []
        for x, y in ring:
            nx = x
            if abs(dx_left) > 1e-7 and ((x <= left_cut) if left_staggered else (abs(x - minx) <= 0.5)):
                nx = x + dx_left
            elif abs(dx_right) > 1e-7 and ((x >= right_cut) if right_staggered else (abs(x - maxx) <= 0.5)):
                nx = x + dx_right
            new_pts.append((nx, y))
        ring = clean_ring(new_pts, min_seg_len=0.05, ccw=True)

    if rounding_mm > 0.0:
        b_minx, _, b_maxx, _ = polygon_bounds(ring)
        width = b_maxx - b_minx
        target_w = round(width / rounding_mm) * rounding_mm
        delta = target_w - width
        if abs(delta) > 1e-6:
            adjust_left = start_tan > end_tan and start_tan > 1e-6
            ring = [
                (
                    (x - delta if (adjust_left and abs(x - b_minx) < 0.5) else (x + delta if ((not adjust_left) and abs(x - b_maxx) < 0.5) else x)),
                    y,
                )
                for x, y in ring
            ]

    return ring


def _transform_and_normalize_plate(
    ring: Ring,
    face: str,
    dstv: DSTVFile,
    cfg: ChesonConfig,
    mirror_x: bool = False,
    mirror_y: bool = False,
    delta_l_right: float = 0.0,
    mid_x_env: Optional[float] = None,
) -> tuple[Ring, list[DSTVHole], list[Ring], list[DSTVMarking], tuple[float, float, float, float]]:
    orig_bounds = polygon_bounds(ring)
    cx = 0.5 * (orig_bounds[0] + orig_bounds[2])
    cy = 0.5 * (orig_bounds[1] + orig_bounds[3])

    def _shift_x(x: float) -> float:
        if mid_x_env is not None and abs(delta_l_right) > 1e-7 and x > mid_x_env:
            return x + delta_l_right
        return x

    def _m_pt(x: float, y: float) -> tuple[float, float]:
        return ((2.0 * cx - x) if mirror_x else x, (2.0 * cy - y) if mirror_y else y)

    mapped_inners: list[Ring] = []
    for ik in dstv.ik_contours.get(face, []):
        ik_ring = contour_to_ring(ik, cfg)
        if len(ik_ring) < 3:
            continue
        ik_ring = [(_shift_x(ix), iy) for ix, iy in ik_ring]
        icx = sum(p[0] for p in ik_ring) / len(ik_ring)
        icy = sum(p[1] for p in ik_ring) / len(ik_ring)
        if point_in_polygon(icx, icy, ring, tol=0.5):
            mapped_inners.append(mirror_ring(ik_ring, cx, cy, mirror_x, mirror_y))

    mapped_holes: list[DSTVHole] = []
    for hole in dstv.holes.get(face, []):
        hx = _shift_x(hole.x)
        hy = hole.y
        is_slot = hole.hole_type == "l" and (abs(hole.slot_length) > 0.05 or abs(hole.slot_width) > 0.05)
        if is_slot:
            rad = math.radians(hole.slot_angle)
            ux, uy = math.cos(rad), math.sin(rad)
            vx, vy = -uy, ux
            dx_s, dy_s = float(hole.slot_length), float(hole.slot_width)
            c1 = (hx, hy)
            mid_check = (hx + 0.5 * (dx_s * ux + dy_s * vx), hy + 0.5 * (dx_s * uy + dy_s * vy))
            if not (
                point_in_polygon(c1[0], c1[1], ring, tol=0.05)
                or point_in_polygon(mid_check[0], mid_check[1], ring, tol=0.05)
            ):
                continue

            if abs(dx_s) > 0.05 and abs(dy_s) > 0.05:
                k0 = c1
                k1 = (c1[0] + dx_s * ux, c1[1] + dx_s * uy)
                k3 = (c1[0] + dy_s * vx, c1[1] + dy_s * vy)
                mk0 = _m_pt(*k0)
                mk1 = _m_pt(*k1)
                mk3 = _m_pt(*k3)
                nx, ny = mk0
                if bool(mirror_x) != bool(mirror_y):
                    out_len = math.hypot(mk3[0] - mk0[0], mk3[1] - mk0[1])
                    out_wid = math.hypot(mk1[0] - mk0[0], mk1[1] - mk0[1])
                    ang_out = math.degrees(math.atan2(mk3[1] - mk0[1], mk3[0] - mk0[0])) % 360.0
                else:
                    out_len = math.hypot(mk1[0] - mk0[0], mk1[1] - mk0[1])
                    out_wid = math.hypot(mk3[0] - mk0[0], mk3[1] - mk0[1])
                    ang_out = math.degrees(math.atan2(mk1[1] - mk0[1], mk1[0] - mk0[0])) % 360.0
            else:
                c2 = (hx + dx_s * ux + dy_s * vx, hy + dx_s * uy + dy_s * vy)
                mc1 = _m_pt(*c1)
                mc2 = _m_pt(*c2)
                nx, ny = mc1
                out_len = math.hypot(mc2[0] - mc1[0], mc2[1] - mc1[1])
                out_wid = 0.0
                ang_out = math.degrees(math.atan2(mc2[1] - mc1[1], mc2[0] - mc1[0])) % 360.0

            mapped_holes.append(
                DSTVHole(
                    face=hole.face,
                    x=nx,
                    y=ny,
                    diameter=hole.diameter,
                    depth=hole.depth,
                    hole_type=hole.hole_type,
                    slot_length=out_len,
                    slot_width=out_wid,
                    slot_angle=ang_out,
                    x_flag=hole.x_flag,
                    y_flag=hole.y_flag,
                )
            )
        elif point_in_polygon(hx, hy, ring, tol=0.05):
            nx, ny = _m_pt(hx, hy)
            ang = hole.slot_angle
            if mirror_x:
                ang = 180.0 - ang
            if mirror_y:
                ang = -ang
            mapped_holes.append(
                DSTVHole(
                    face=hole.face,
                    x=nx,
                    y=ny,
                    diameter=hole.diameter,
                    depth=hole.depth,
                    hole_type=hole.hole_type,
                    slot_length=hole.slot_length,
                    slot_width=hole.slot_width,
                    slot_angle=ang % 360.0,
                    x_flag=hole.x_flag,
                    y_flag=hole.y_flag,
                )
            )

    mapped_markings: list[DSTVMarking] = []
    for mark in dstv.markings.get(face, []):
        mx = _shift_x(mark.x)
        if point_in_polygon(mx, mark.y, ring, tol=0.05):
            nx, ny = _m_pt(mx, mark.y)
            mapped_markings.append(
                DSTVMarking(
                    face=mark.face,
                    x=nx,
                    y=ny,
                    angle=mark.angle,
                    height=mark.height,
                    text=mark.text,
                    x_flag=mark.x_flag,
                    y_flag=mark.y_flag,
                )
            )

    t_ring = mirror_ring(ring, cx, cy, mirror_x, mirror_y)
    minx, miny, _, _ = polygon_bounds(t_ring)

    shifted_ext = clean_ring(
        translate_ring(t_ring, -minx, -miny),
        min_seg_len=cfg.min_segment_length_mm,
        ccw=True,
    )
    shifted_inners = [
        clean_ring(translate_ring(ir, -minx, -miny), min_seg_len=cfg.min_segment_length_mm, ccw=False)
        for ir in mapped_inners
    ]
    shifted_holes = [
        DSTVHole(
            face=h.face,
            x=h.x - minx,
            y=h.y - miny,
            diameter=h.diameter,
            depth=h.depth,
            hole_type=h.hole_type,
            slot_length=h.slot_length,
            slot_width=h.slot_width,
            slot_angle=h.slot_angle,
            x_flag=h.x_flag,
            y_flag=h.y_flag,
        )
        for h in mapped_holes
    ]
    shifted_markings = [
        DSTVMarking(
            face=m.face,
            x=m.x - minx,
            y=m.y - miny,
            angle=m.angle,
            height=m.height,
            text=m.text,
            x_flag=m.x_flag,
            y_flag=m.y_flag,
        )
        for m in mapped_markings
    ]
    return shifted_ext, shifted_holes, shifted_inners, shifted_markings, orig_bounds


def _clip_single_band_web(
    ring: Ring,
    H: float,
    tw: float,
    tf_top: float,
    tf_bot: float,
    cheson_type: str,
    talpa_side: str,
    top_over: bool,
    bot_over: bool,
    cfg: ChesonConfig,
) -> Ring:
    _, miny, _, maxy = polygon_bounds(ring)
    curr_h = maxy - miny

    if cfg.override_web_height_mm is not None and cfg.override_web_height_mm > 0.0:
        target_h = float(cfg.override_web_height_mm)
        if abs(curr_h - target_h) <= 0.001:
            return ring
        if talpa_side == "bottom":
            y_lo = miny if miny >= tf_bot - 1.0 else max(miny, tf_bot)
            y_hi = min(maxy, y_lo + target_h)
            if y_hi - y_lo < target_h - 0.5:
                y_lo = max(miny, y_hi - target_h)
        else:
            y_lo = miny + max(0.0, 0.5 * (curr_h - target_h))
            y_hi = y_lo + target_h
        return clip_ring_y_strip(ring, y_lo, y_hi, min_seg_len=cfg.min_segment_length_mm)

    if cheson_type == "cheson_cu_talpa":
        if talpa_side == "top":
            target_h = H - tf_top - tw
            if abs(curr_h - target_h) <= 0.5:
                return ring
            y_hi = maxy if maxy <= H - tf_top + 1.0 else (H - tf_top)
            y_lo = y_hi - target_h
        else:
            # Standard Tip 1: wide bottom talpă + narrow top capac between webs
            target_h = H - tf_bot - tw
            if abs(curr_h - target_h) <= 0.5:
                return ring
            if miny <= 1.0 and curr_h >= H - 1.0:
                y_lo = tf_bot
                y_hi = H - tw
            elif miny >= tf_bot - 1.0:
                y_lo = miny
                y_hi = miny + target_h
            else:
                # Tekla exported [0, H - tf_bot]
                y_lo = miny
                y_hi = miny + target_h
        return clip_ring_y_strip(ring, y_lo, y_hi, min_seg_len=cfg.min_segment_length_mm)

    if cheson_type == "cheson_fara_talpa":
        if cfg.flange_top_mode == "over_webs" and cfg.flange_bot_mode == "over_webs":
            return ring
        target_h = H - 2.0 * tw
        if abs(curr_h - target_h) <= 0.5:
            return ring
        if miny <= 1.0 and maxy >= H - 1.0:
            y_lo = tw
            y_hi = H - tw
        else:
            y_lo = miny + tw
            y_hi = maxy - tw
        return clip_ring_y_strip(ring, y_lo, y_hi, min_seg_len=cfg.min_segment_length_mm)

    # Fallback nominal parametric
    top_recess = resolve_recess_mm(cfg.web_top_recess_mm, tw)
    bot_recess = resolve_recess_mm(cfg.web_bot_recess_mm, tw)
    y_lo = miny
    y_hi = maxy
    if maxy >= H - 1.0:
        if top_over:
            y_hi = min(y_hi, H - tf_top)
        elif top_recess > 0.0:
            y_hi = min(y_hi, H - top_recess)
    if miny <= 1.0:
        if bot_over:
            y_lo = max(y_lo, tf_bot)
        elif bot_recess > 0.0:
            y_lo = max(y_lo, bot_recess)
    if y_lo > miny + 1e-6 or y_hi < maxy - 1e-6:
        return clip_ring_y_strip(ring, y_lo, y_hi, min_seg_len=cfg.min_segment_length_mm)
    return ring


def _detect_tip1_recessed_web_faces(
    locked_rings: dict[str, list[Ring]],
    spec: CompoundProfileSpec,
    detected_type: str,
    talpa_side: str,
) -> tuple[set[str], str, tuple[float, float]]:
    """Detect Tip 1 overhanging-talpă faces (v/h) whose DSTV AK only shows the talpă edge strip."""
    if detected_type not in ("cheson_cu_talpa", "nominal_parametric"):
        return set(), "", (0.0, spec.total_width)

    if talpa_side not in ("bottom", "top"):
        u_rings = locked_rings.get("u", [])
        o_rings = locked_rings.get("o", [])
        dy_u = (max(polygon_bounds(r)[3] for r in u_rings) - min(polygon_bounds(r)[1] for r in u_rings)) if u_rings else spec.b_bot
        dy_o = (max(polygon_bounds(r)[3] for r in o_rings) - min(polygon_bounds(r)[1] for r in o_rings)) if o_rings else spec.b_top
        if dy_u > dy_o + 2.0 * spec.tw + 1.0:
            talpa_side = "bottom"
        elif dy_o > dy_u + 2.0 * spec.tw + 1.0:
            talpa_side = "top"
        else:
            return set(), "", (0.0, spec.total_width)

    talpa_face = "u" if talpa_side == "bottom" else "o"
    narrow_face = "o" if talpa_side == "bottom" else "u"
    tf_talpa = spec.tf_bot if talpa_side == "bottom" else spec.tf_top

    t_rings = locked_rings.get(talpa_face, [])
    n_rings = locked_rings.get(narrow_face, [])
    if not t_rings or not n_rings:
        return set(), narrow_face, (0.0, spec.total_width)

    y_t_lo = min(polygon_bounds(r)[1] for r in t_rings)
    y_t_hi = max(polygon_bounds(r)[3] for r in t_rings)
    y_n_lo = min(polygon_bounds(r)[1] for r in n_rings)
    y_n_hi = max(polygon_bounds(r)[3] for r in n_rings)

    recessed: set[str] = set()
    h_rings = locked_rings.get("h", [])
    h_max_h = max((polygon_bounds(r)[3] - polygon_bounds(r)[1] for r in h_rings), default=0.0)
    if (y_n_lo > y_t_lo + 1.0) and (not h_rings or h_max_h <= tf_talpa + 1.0):
        recessed.add("h")

    v_rings = locked_rings.get("v", [])
    v_max_h = max((polygon_bounds(r)[3] - polygon_bounds(r)[1] for r in v_rings), default=0.0)
    if (y_n_hi < y_t_hi - 1.0) and (not v_rings or v_max_h <= tf_talpa + 1.0):
        recessed.add("v")

    return recessed, narrow_face, (y_n_lo, y_n_hi)


def _synthesize_recessed_web_rings(
    face: str,
    narrow_rings: list[Ring],
    y_n_lo: float,
    y_n_hi: float,
    spec: CompoundProfileSpec,
    talpa_side: str,
    cfg: ChesonConfig,
) -> list[Ring]:
    """Reconstruct recessed web contour(s) from the narrow flange outer box (`o`/`u`)."""
    H = spec.total_height
    tw = spec.tw
    eff_talpa = talpa_side if talpa_side in ("bottom", "top") else (
        "bottom" if spec.b_bot >= spec.b_top else "top"
    )
    if cfg.override_web_height_mm is not None and cfg.override_web_height_mm > 0.0:
        target_h = float(cfg.override_web_height_mm)
        if eff_talpa == "bottom":
            z_lo = spec.tf_bot
            z_hi = z_lo + target_h
        else:
            z_hi = H - spec.tf_top
            z_lo = z_hi - target_h
    else:
        if eff_talpa == "bottom":
            z_lo = spec.tf_bot
            z_hi = H - tw
        else:
            z_lo = tw
            z_hi = H - spec.tf_top

    if z_hi - z_lo <= 1.0:
        return []

    y_outer = y_n_hi if face == "v" else y_n_lo
    y_inner = (y_n_hi - tw) if face == "v" else (y_n_lo + tw)

    synthesized: list[Ring] = []
    for nr in narrow_rings:
        if len(nr) < 3 or abs(polygon_area(nr)) <= 0.1:
            continue
        sp_out = _x_span_at_y(nr, y_outer)
        sp_in = _x_span_at_y(nr, y_inner)
        if sp_out is None and sp_in is None:
            continue
        xs_lo = [sp[0] for sp in (sp_out, sp_in) if sp is not None]
        xs_hi = [sp[1] for sp in (sp_out, sp_in) if sp is not None]
        x0 = min(xs_lo)
        x1 = max(xs_hi)
        if x1 - x0 <= 1.0:
            continue
        rect: Ring = [(x0, z_lo), (x1, z_lo), (x1, z_hi), (x0, z_hi)]
        rect = _apply_length_offset(rect, cfg.length_offset_mm)
        if len(rect) >= 3 and abs(polygon_area(rect)) > 0.1:
            synthesized.append(rect)
    return synthesized


def _narrow_face_web_x_bounds(
    narrow_rings: list[Ring],
    face: str,
    y_n_lo: float,
    y_n_hi: float,
    tw: float,
) -> Optional[tuple[float, float]]:
    if not narrow_rings:
        return None
    y_outer = y_n_hi if face == "v" else y_n_lo
    y_inner = (y_n_hi - tw) if face == "v" else (y_n_lo + tw)
    xs_lo: list[float] = []
    xs_hi: list[float] = []
    for nr in narrow_rings:
        for yt in (y_outer, y_inner):
            sp = _x_span_at_y(nr, yt)
            if sp is not None:
                xs_lo.append(sp[0])
                xs_hi.append(sp[1])
    return (min(xs_lo), max(xs_hi)) if (xs_lo and xs_hi) else None


def verify_cheson_decomposition(
    dstv: DSTVFile,
    plates: list[ExtractedPlate],
    l_nc1_raw: float,
    l_gabarit_target: float,
    l_source_label: str,
    l_assembled: float,
    detected_type: str,
    talpa_side: str,
    spec: CompoundProfileSpec,
    cfg: Optional[ChesonConfig] = None,
) -> VerificationReport:
    """Closed-loop geometric audit verifying <= 0.001 mm tolerance on gabarit & flange-web contact."""
    cfg = cfg or ChesonConfig()
    H_target = spec.total_height
    W_target = spec.total_width
    tw = spec.tw

    flanges = [p for p in plates if p.role in ("flange_top", "flange_bot", "talpa", "capac")]
    webs = [p for p in plates if p.role == "web"]

    # Check minimum required structural plates for compound box beams
    min_webs = 4 if detected_type == "cheson_dublu" else 2
    min_flanges = 2
    missing_components = (len(webs) < min_webs) or (len(flanges) < min_flanges)

    # Check assembled width
    if detected_type == "cheson_cu_talpa":
        talpa_p = [p for p in plates if p.role == "talpa"]
        capac_p = [p for p in plates if p.role == "capac"]
        w_talpa = (talpa_p[0].bounds[3] - talpa_p[0].bounds[1]) if talpa_p else 0.0
        w_capac = (capac_p[0].bounds[3] - capac_p[0].bounds[1] + 2.0 * tw) if capac_p else 0.0
        w_assembled = max(w_talpa, w_capac)
    elif detected_type in ("cheson_fara_talpa", "cheson_dublu"):
        w_vals = []
        for fp in flanges:
            fw = fp.bounds[3] - fp.bounds[1]
            # If flange sits between webs, assembled width = fw + 2*tw
            if abs(fw - (W_target - 2.0 * tw)) <= 1.0:
                w_vals.append(fw + 2.0 * tw)
            else:
                w_vals.append(fw)
        w_assembled = max(w_vals) if w_vals else W_target
    else:
        w_assembled = max((fp.bounds[3] - fp.bounds[1]) for fp in flanges) if flanges else W_target

    # Check assembled height across ALL webs (never mask a bogus/missing web)
    if detected_type == "cheson_dublu" and len(webs) >= 2:
        unique_web_hs = sorted({round(wp.bounds[3] - wp.bounds[1], 4) for wp in webs}, reverse=True)
        # Find gap between upper and lower web bands from raw NC1
        gap_mm = 10.0
        for f in ("v", "h"):
            cs = dstv.ak_contours.get(f, [])
            if len(cs) >= 2:
                ys_pair = sorted([polygon_bounds(contour_to_ring(c))[1:4:2] for c in cs], key=lambda t: t[0])
                gap_mm = max(0.0, ys_pair[1][0] - ys_pair[0][1])
                break
        if len(unique_web_hs) >= 2:
            h_assembled = unique_web_hs[0] + unique_web_hs[1] + gap_mm + 2.0 * tw
        elif unique_web_hs:
            h_assembled = 2.0 * unique_web_hs[0] + gap_mm + 2.0 * tw
        else:
            h_assembled = 0.0
    elif detected_type == "cheson_cu_talpa" and webs:
        tf_t = spec.tf_top if talpa_side == "top" else spec.tf_bot
        h_candidates = [(wp.bounds[3] - wp.bounds[1]) + tf_t + tw for wp in webs]
        h_assembled = max(h_candidates, key=lambda h_v: abs(h_v - H_target))
    elif detected_type == "cheson_fara_talpa" and webs:
        is_over_w = bool(flanges and abs((flanges[0].bounds[3] - flanges[0].bounds[1]) - W_target) <= 0.5)
        h_candidates = [
            ((wp.bounds[3] - wp.bounds[1]) + (spec.tf_top + spec.tf_bot if is_over_w else 2.0 * tw))
            for wp in webs
        ]
        h_assembled = max(h_candidates, key=lambda h_v: abs(h_v - H_target))
    else:
        h_assembled = 0.0 if missing_components else H_target

    # Check global-coordinate seam contact between adjacent capac flanges and webs
    contact_err = 0.0
    capac_flanges = [p for p in plates if p.role == "capac"]
    if not capac_flanges:
        capac_flanges = [p for p in plates if p.role in ("flange_top", "talpa")]
    for ref_flange in capac_flanges:
        mir_y = cfg.mirror_flange_o_y if ref_flange.face == "o" else cfg.mirror_flange_u_y
        f_glob = _untransform_dxf_ring(
            ref_flange.exterior, ref_flange.source_bounds, mirror_x=False, mirror_y=mir_y
        )
        _, fy_lo, _, fy_hi = ref_flange.source_bounds
        for wp in webs:
            if wp.removed_residuals != 0:
                continue
            wz0, wz1 = wp.source_bounds[1], wp.source_bounds[3]
            if detected_type == "cheson_dublu":
                wz_mid = 0.5 * (wz0 + wz1)
                if ref_flange.face == "o" and wz_mid < 0.5 * H_target:
                    continue
                if ref_flange.face == "u" and wz_mid > 0.5 * H_target:
                    continue
            mir_x = cfg.mirror_web_h_x if wp.face == "h" else False
            w_glob = _untransform_dxf_ring(wp.exterior, wp.source_bounds, mirror_x=mir_x, mirror_y=False)
            contact_y = fy_lo if wp.face == "h" else fy_hi
            f_span = _x_span_at_y(f_glob, contact_y)
            w_z_ref = wz1 if ref_flange.face == "o" else wz0
            w_span = _x_span_at_y(w_glob, w_z_ref)
            if f_span is not None and w_span is not None:
                gap_l = max(0.0, w_span[0] - f_span[0])
                gap_r = max(0.0, f_span[1] - w_span[1])
                contact_err = max(contact_err, gap_l, gap_r)

    err_l = abs(l_assembled - l_gabarit_target)
    err_h = abs(h_assembled - H_target)
    err_w = abs(w_assembled - W_target)
    max_err = max(err_l, err_h, err_w)

    return VerificationReport(
        piece_id=dstv.header.piece_id,
        detected_type=detected_type,
        detected_type_label=CHESON_TYPE_LABELS.get(detected_type, detected_type),
        l_nc1_raw_mm=round(l_nc1_raw, 4),
        l_gabarit_target_mm=round(l_gabarit_target, 4),
        l_source_label=l_source_label,
        l_assembled_mm=round(l_assembled, 4),
        h_target_mm=round(H_target, 4),
        h_assembled_mm=round(h_assembled, 4),
        w_target_mm=round(W_target, 4),
        w_assembled_mm=round(w_assembled, 4),
        max_gabarit_error_mm=round(max_err, 6),
        flange_web_contact_error_mm=round(contact_err, 6),
        passed=(not missing_components) and (max_err <= 0.001 + 1e-9) and (contact_err <= 0.05 + 1e-9),
    )


def decompose_compound_nc1(
    dstv: DSTVFile,
    cfg: Optional[ChesonConfig] = None,
    return_verification: bool = False,
):
    cfg = cfg or ChesonConfig()
    header = dstv.header
    spec = parse_compound_profile(header.profile_str, header)
    spec = _align_spec_flanges_to_faces(dstv, spec, cfg)
    piece_clean = clean_piece_name(header.piece_id, lowercase=cfg.lowercase_piece_prefix)
    detected_type, talpa_side = detect_cheson_type(dstv, spec, cfg)

    H = spec.total_height
    B = spec.total_width
    tw = spec.tw

    # 1. Convert raw AK rings & apply exact gabarit X lock
    raw_rings: dict[str, list[Ring]] = {}
    for face in ("o", "u", "v", "h"):
        raw_rings[face] = []
        for c in dstv.ak_contours.get(face, []):
            r = contour_to_ring(c, cfg)
            if len(r) >= 3 and abs(polygon_area(r)) > 0.1:
                raw_rings[face].append(r)

    x_min_env, x_max_env, l_nc1_raw = _compute_raw_x_envelope(raw_rings)
    l_gabarit, l_source_label = _resolve_gabarit_length(l_nc1_raw, cfg)

    if cfg.exact_gabarit_lock:
        max_delta = 100.0 if (cfg.target_length_mm is not None and cfg.target_length_mm > 0.0) else 20.0
        locked_rings = _apply_gabarit_x_lock_to_rings(
            raw_rings, x_min_env, l_nc1_raw, l_gabarit, max_delta_mm=max_delta
        )
    else:
        locked_rings = raw_rings

    x_min_locked, x_max_locked, l_assembled_env = _compute_raw_x_envelope(locked_rings)
    mid_x_raw = x_min_env + 0.5 * l_nc1_raw
    applied_delta_l = (l_assembled_env - l_nc1_raw) if cfg.exact_gabarit_lock else 0.0
    start_tan, end_tan = _estimate_end_skew_angles(dstv, locked_rings)

    # Detect Tip 1 overhanging-talpă faces whose AK v/h only captures the talpă edge strip
    recessed_faces, narrow_face, (y_n_lo, y_n_hi) = _detect_tip1_recessed_web_faces(
        locked_rings, spec, detected_type, talpa_side
    )

    # Determine flange clipping modes & strict roles (talpa='T' vs capac='C') per detected cheson type
    dy_o_raw = _face_y_span(dstv, "o", cfg) or spec.b_top
    dy_u_raw = _face_y_span(dstv, "u", cfg) or spec.b_bot
    if detected_type == "cheson_cu_talpa":
        top_over = (talpa_side == "top")
        bot_over = (talpa_side == "bottom")
        o_suffix = "T" if talpa_side == "top" else cfg.flange_letter
        u_suffix = "T" if talpa_side == "bottom" else cfg.flange_letter
        o_role = "talpa" if talpa_side == "top" else "capac"
        u_role = "talpa" if talpa_side == "bottom" else "capac"
    elif detected_type == "cheson_dublu":
        top_over = dy_o_raw > dy_u_raw + 2.0 * tw + 1.0
        bot_over = dy_u_raw > dy_o_raw + 2.0 * tw + 1.0
        o_role = "talpa" if top_over else "capac"
        u_role = "talpa" if bot_over else "capac"
        o_suffix = "T" if top_over else (cfg.flange_letter if bot_over else f"{cfg.flange_letter}1")
        u_suffix = "T" if bot_over else (cfg.flange_letter if top_over else f"{cfg.flange_letter}2")
    elif detected_type == "cheson_fara_talpa":
        top_over = (cfg.flange_top_mode == "over_webs")
        bot_over = (cfg.flange_bot_mode == "over_webs")
        o_suffix = f"{cfg.flange_letter}1"
        u_suffix = f"{cfg.flange_letter}2"
        o_role = "capac"
        u_role = "capac"
    else:
        top_over = _resolve_flange_over_webs(cfg.flange_top_mode, spec.top_over_webs)
        bot_over = _resolve_flange_over_webs(cfg.flange_bot_mode, spec.bot_over_webs)
        o_is_wide = dy_o_raw > dy_u_raw + 2.0 * tw + 1.0
        u_is_wide = dy_u_raw > dy_o_raw + 2.0 * tw + 1.0
        o_role = "talpa" if o_is_wide else "capac"
        u_role = "talpa" if u_is_wide else "capac"
        o_suffix = "T" if o_is_wide else f"{cfg.flange_letter}1"
        u_suffix = "T" if u_is_wide else f"{cfg.flange_letter}2"

    plates: list[ExtractedPlate] = []
    clipped_flanges_by_face: dict[str, list[Ring]] = {"o": [], "u": []}

    # 2. Flanges: o and u
    flange_specs = [
        ("o", o_role, spec.tf_top, top_over, cfg.mirror_flange_o_y, o_suffix),
        ("u", u_role, spec.tf_bot, bot_over, cfg.mirror_flange_u_y, u_suffix),
    ]
    for face, role, thick, is_over, mirror_y, suffix in flange_specs:
        flange_part_idx = 0
        for ring in locked_rings.get(face, []):
            _, f_miny, _, f_maxy = polygon_bounds(ring)
            f_span = f_maxy - f_miny
            if (not is_over) and f_span > 2.0 * tw + 1.0:
                # Clip symmetric tw from both sides of the flange face's own bounds
                ring = clip_ring_y_strip(
                    ring,
                    f_miny + tw,
                    f_maxy - tw,
                    min_seg_len=cfg.min_segment_length_mm,
                )
            if len(ring) < 3 or abs(polygon_area(ring)) <= 0.1:
                continue
            ring = _apply_length_offset(ring, cfg.length_offset_mm)
            if cfg.strip_residual_fragments:
                sub_rings, rem_cnt = remove_residual_fragments(
                    ring, thick, min_seg_len=cfg.min_segment_length_mm
                )
            else:
                sub_rings, rem_cnt = [ring], 0

            for sub_ring in sub_rings:
                if len(sub_ring) < 3 or abs(polygon_area(sub_ring)) <= 0.1:
                    continue
                clipped_flanges_by_face[face].append(sub_ring)
                plate_suffix = suffix if flange_part_idx == 0 else f"{suffix}_{flange_part_idx + 1}"
                flange_part_idx += 1
                s_ext, s_holes, s_inners, s_marks, orig_b = _transform_and_normalize_plate(
                    sub_ring,
                    face=face,
                    dstv=dstv,
                    cfg=cfg,
                    mirror_x=False,
                    mirror_y=mirror_y,
                    delta_l_right=applied_delta_l,
                    mid_x_env=mid_x_raw,
                )
                t_int = int(round(thick))
                plates.append(
                    ExtractedPlate(
                        name=f"{piece_clean}{plate_suffix}",
                        folder=f"T{t_int}",
                        role=role,
                        face=face,
                        thickness=thick,
                        piece_id=header.piece_id,
                        piece_clean=piece_clean,
                        material=header.material,
                        quantity=header.quantity,
                        exterior=s_ext,
                        holes=s_holes,
                        inner_contours=s_inners,
                        markings=s_marks,
                        source_bounds=orig_b,
                        removed_residuals=rem_cnt,
                        role_code=plate_suffix,
                    )
                )

    # Reference flange contact intersections for skewed ends (prefer top capac 'o' or 'u')
    ref_flange_face = "o" if (clipped_flanges_by_face["o"] and not top_over) else (
        "u" if (clipped_flanges_by_face["u"] and not bot_over) else "o"
    )
    ref_rings = clipped_flanges_by_face.get(ref_flange_face, [])
    left_contact_pair = None
    right_contact_pair = None
    if ref_rings:
        _, ref_ylo, _, ref_yhi = polygon_bounds(ref_rings[0])
        left_contact_pair, right_contact_pair = _flange_skew_x_contacts(ref_rings, ref_ylo, ref_yhi)

    # Determine which web face (v or h) is the longer vs shorter side at skewed ends.
    # For recessed faces on Tip 1 overhanging chesons, derive bounds from narrow_face instead of bogus talpă strips.
    narrow_rings_locked = locked_rings.get(narrow_face, []) if narrow_face else []
    if "v" in recessed_faces and narrow_rings_locked:
        vb = _narrow_face_web_x_bounds(narrow_rings_locked, "v", y_n_lo, y_n_hi, tw)
        v_xmin, v_xmax = vb if vb else (0.0, 0.0)
    else:
        v_xmin = min((polygon_bounds(r)[0] for r in locked_rings.get("v", [])), default=0.0)
        v_xmax = max((polygon_bounds(r)[2] for r in locked_rings.get("v", [])), default=0.0)

    if "h" in recessed_faces and narrow_rings_locked:
        hb = _narrow_face_web_x_bounds(narrow_rings_locked, "h", y_n_lo, y_n_hi, tw)
        h_xmin, h_xmax = hb if hb else (0.0, 0.0)
    else:
        h_xmin = min((polygon_bounds(r)[0] for r in locked_rings.get("h", [])), default=0.0)
        h_xmax = max((polygon_bounds(r)[2] for r in locked_rings.get("h", [])), default=0.0)

    eff_start_tan = start_tan if abs(v_xmin - h_xmin) > 0.5 else 0.0
    eff_end_tan = end_tan if abs(v_xmax - h_xmax) > 0.5 else 0.0

    # 3. Webs: v (front) and h (back)
    raw_webs: list[tuple[str, Ring, int]] = []
    is_multi_band = _has_vertical_web_gap(dstv.ak_contours.get("v", []), cfg) or _has_vertical_web_gap(
        dstv.ak_contours.get("h", []), cfg
    )
    top_recess = resolve_recess_mm(cfg.web_top_recess_mm, tw)
    bot_recess = resolve_recess_mm(cfg.web_bot_recess_mm, tw)

    for face in ("v", "h"):
        if face in recessed_faces:
            synth_rings = _synthesize_recessed_web_rings(
                face=face,
                narrow_rings=narrow_rings_locked,
                y_n_lo=y_n_lo,
                y_n_hi=y_n_hi,
                spec=spec,
                talpa_side=talpa_side,
                cfg=cfg,
            )
            for s_ring in synth_rings:
                if cfg.strip_residual_fragments:
                    sub_rings, rem_cnt = remove_residual_fragments(
                        s_ring, tw, min_seg_len=cfg.min_segment_length_mm
                    )
                else:
                    sub_rings, rem_cnt = [s_ring], 0
                for sr in sub_rings:
                    if len(sr) >= 3 and abs(polygon_area(sr)) > 0.1:
                        raw_webs.append((face, sr, rem_cnt))
            continue

        ref_orig_rings = locked_rings.get(ref_flange_face, [])
        wall_lx: Optional[float] = None
        wall_rx: Optional[float] = None
        if ref_orig_rings:
            ror = ref_orig_rings[0]
            _, ro_ylo, _, ro_yhi = polygon_bounds(ror)
            yw_out = ro_ylo if face == "h" else ro_yhi
            yw_in = (ro_ylo + tw) if face == "h" else (ro_yhi - tw)
            sp_w_o = _x_span_at_y(ror, yw_out)
            sp_w_i = _x_span_at_y(ror, yw_in)
            w_xs_l = [sp[0] for sp in (sp_w_o, sp_w_i) if sp is not None]
            w_xs_r = [sp[1] for sp in (sp_w_o, sp_w_i) if sp is not None]
            if w_xs_l:
                wall_lx = min(w_xs_l)
            if w_xs_r:
                wall_rx = max(w_xs_r)

        is_long_left: Optional[bool] = None
        sync_lx: Optional[float] = None
        if eff_start_tan > 1e-6:
            is_long_left = (face == "v" and v_xmin < h_xmin) or (face == "h" and h_xmin < v_xmin)
            if is_long_left:
                if wall_lx is not None:
                    sync_lx = wall_lx
                else:
                    sync_lx = min(v_xmin, h_xmin) if recessed_faces else x_min_locked
            elif left_contact_pair is not None:
                sync_lx = left_contact_pair[1]

        is_long_right: Optional[bool] = None
        sync_rx: Optional[float] = None
        if eff_end_tan > 1e-6:
            is_long_right = (face == "v" and v_xmax > h_xmax) or (face == "h" and h_xmax > v_xmax)
            if is_long_right:
                if wall_rx is not None:
                    sync_rx = wall_rx
                else:
                    sync_rx = max(v_xmax, h_xmax) if recessed_faces else x_max_locked
            elif right_contact_pair is not None:
                sync_rx = right_contact_pair[0]

        for ring in locked_rings.get(face, []):
            if is_multi_band:
                _, miny, _, maxy = polygon_bounds(ring)
                y_lo, y_hi = miny, maxy
                if maxy >= H - 1.0:
                    if top_over:
                        y_hi = min(y_hi, H - spec.tf_top)
                    elif top_recess > 0.0:
                        y_hi = min(y_hi, H - top_recess)
                if miny <= 1.0:
                    if bot_over:
                        y_lo = max(y_lo, spec.tf_bot)
                    elif bot_recess > 0.0:
                        y_lo = max(y_lo, bot_recess)
                if y_lo > miny + 1e-6 or y_hi < maxy - 1e-6:
                    ring = clip_ring_y_strip(ring, y_lo, y_hi, min_seg_len=cfg.min_segment_length_mm)
            else:
                ring = _clip_single_band_web(
                    ring,
                    H=H,
                    tw=tw,
                    tf_top=spec.tf_top,
                    tf_bot=spec.tf_bot,
                    cheson_type=detected_type,
                    talpa_side=talpa_side,
                    top_over=top_over,
                    bot_over=bot_over,
                    cfg=cfg,
                )

            if len(ring) < 3 or abs(polygon_area(ring)) <= 0.1:
                continue

            ring = _adjust_web_skew_and_rounding(
                ring,
                face=face,
                tw=tw,
                start_tan=eff_start_tan,
                end_tan=eff_end_tan,
                skew_plane=cfg.web_skew_plane,
                rounding_mm=cfg.web_length_rounding_mm,
                sync_left_x=sync_lx if cfg.exact_gabarit_lock else None,
                sync_right_x=sync_rx if cfg.exact_gabarit_lock else None,
                is_long_left=is_long_left,
                is_long_right=is_long_right,
            )
            ring = _apply_length_offset(ring, cfg.length_offset_mm)
            if len(ring) >= 3 and abs(polygon_area(ring)) > 0.1:
                if cfg.strip_residual_fragments:
                    sub_rings, rem_cnt = remove_residual_fragments(
                        ring, tw, min_seg_len=cfg.min_segment_length_mm
                    )
                else:
                    sub_rings, rem_cnt = [ring], 0
                for sr in sub_rings:
                    if len(sr) >= 3 and abs(polygon_area(sr)) > 0.1:
                        raw_webs.append((face, sr, rem_cnt))

    if (cfg.web_ordering or "").lower() == "v_then_h_upper_first":
        raw_webs.sort(
            key=lambda item: (
                0 if item[0] == "v" else 1,
                -round(0.5 * (polygon_bounds(item[1])[1] + polygon_bounds(item[1])[3]), 1),
                round(polygon_bounds(item[1])[0], 1),
            )
        )
    else:
        raw_webs.sort(
            key=lambda item: (
                -round(0.5 * (polygon_bounds(item[1])[1] + polygon_bounds(item[1])[3]), 1),
                0 if item[0] == "h" else 1,
                round(polygon_bounds(item[1])[0], 1),
            )
        )

    tw_int = int(round(tw))
    for w_idx, (face, w_ring, rem_cnt) in enumerate(raw_webs, start=1):
        mirror_x = cfg.mirror_web_h_x if face == "h" else False
        s_ext, s_holes, s_inners, s_marks, orig_b = _transform_and_normalize_plate(
            w_ring,
            face=face,
            dstv=dstv,
            cfg=cfg,
            mirror_x=mirror_x,
            mirror_y=False,
            delta_l_right=applied_delta_l,
            mid_x_env=mid_x_raw,
        )
        plates.append(
            ExtractedPlate(
                name=f"{piece_clean}{cfg.web_letter}{w_idx}",
                folder=f"T{tw_int}",
                role="web",
                face=face,
                thickness=tw,
                piece_id=header.piece_id,
                piece_clean=piece_clean,
                material=header.material,
                quantity=header.quantity,
                exterior=s_ext,
                holes=s_holes,
                inner_contours=s_inners,
                markings=s_marks,
                source_bounds=orig_b,
                removed_residuals=rem_cnt,
                role_code=f"{cfg.web_letter}{w_idx}",
            )
        )

    if not return_verification:
        return plates

    ver = verify_cheson_decomposition(
        dstv=dstv,
        plates=plates,
        l_nc1_raw=l_nc1_raw,
        l_gabarit_target=l_gabarit,
        l_source_label=l_source_label,
        l_assembled=l_assembled_env + cfg.length_offset_mm,
        detected_type=detected_type,
        talpa_side=talpa_side,
        spec=spec,
        cfg=cfg,
    )
    return plates, ver


def extract_single_plate_nc1(
    dstv: DSTVFile,
    cfg: Optional[ChesonConfig] = None,
) -> list[ExtractedPlate]:
    cfg = cfg or ChesonConfig()
    header = dstv.header
    piece_clean = clean_piece_name(header.piece_id, lowercase=cfg.lowercase_piece_prefix)
    thick = header.flange_thickness or header.web_thickness or 10.0
    t_int = int(round(thick))

    primary_face = None
    for f in ("v", "o", "u", "h"):
        if dstv.ak_contours.get(f):
            primary_face = f
            break
    if primary_face is None:
        return []

    plates: list[ExtractedPlate] = []
    contours = dstv.ak_contours[primary_face]
    max_delta = 100.0 if (cfg.target_length_mm is not None and cfg.target_length_mm > 0.0) else 20.0
    part_idx = 0
    for c in contours:
        ring = contour_to_ring(c, cfg)
        if len(ring) < 3 or abs(polygon_area(ring)) <= 0.1:
            continue
        x0, _, l_raw = _compute_raw_x_envelope({primary_face: [ring]})
        l_gab, _ = _resolve_gabarit_length(l_raw, cfg)
        mid_x_raw = x0 + 0.5 * l_raw
        applied_delta = 0.0
        if cfg.exact_gabarit_lock:
            ring = _apply_gabarit_x_lock_to_rings(
                {primary_face: [ring]}, x0, l_raw, l_gab, max_delta_mm=max_delta
            )[primary_face][0]
            _, _, l_after = _compute_raw_x_envelope({primary_face: [ring]})
            applied_delta = l_after - l_raw
        ring = _apply_length_offset(ring, cfg.length_offset_mm)
        if cfg.strip_residual_fragments:
            sub_rings, rem_cnt = remove_residual_fragments(
                ring, thick, min_seg_len=cfg.min_segment_length_mm
            )
        else:
            sub_rings, rem_cnt = [ring], 0

        for sub_ring in sub_rings:
            if len(sub_ring) < 3 or abs(polygon_area(sub_ring)) <= 0.1:
                continue
            s_ext, s_holes, s_inners, s_marks, orig_b = _transform_and_normalize_plate(
                sub_ring,
                face=primary_face,
                dstv=dstv,
                cfg=cfg,
                delta_l_right=applied_delta,
                mid_x_env=mid_x_raw,
            )
            name = piece_clean if (len(contours) == 1 and len(sub_rings) == 1 and part_idx == 0) else f"{piece_clean}_{part_idx + 1}"
            part_idx += 1
            plates.append(
                ExtractedPlate(
                    name=name,
                    folder=f"T{t_int}",
                    role="single_plate",
                    face=primary_face,
                    thickness=thick,
                    piece_id=header.piece_id,
                    piece_clean=piece_clean,
                    material=header.material,
                    quantity=header.quantity,
                    exterior=s_ext,
                    holes=s_holes,
                    inner_contours=s_inners,
                    markings=s_marks,
                    source_bounds=orig_b,
                    removed_residuals=rem_cnt,
                    role_code="P" if (len(contours) == 1 and len(sub_rings) == 1 and part_idx == 1) else f"P{part_idx}",
                )
            )
    return plates


def decompose_nc1(
    dstv: DSTVFile,
    cfg: Optional[ChesonConfig] = None,
    return_verification: bool = False,
):
    spec = parse_compound_profile(dstv.header.profile_str, dstv.header)
    if spec.profile_type in ("HSQ", "BOX", "WQ"):
        return decompose_compound_nc1(dstv, cfg, return_verification=return_verification)
    plates = extract_single_plate_nc1(dstv, cfg)
    if not return_verification:
        return plates
    l_raw = max((p.bounds[2] - p.bounds[0]) for p in plates) if plates else 0.0
    ver = VerificationReport(
        piece_id=dstv.header.piece_id,
        detected_type="single_plate",
        detected_type_label="Placă individuală .nc1",
        l_nc1_raw_mm=round(l_raw, 4),
        l_gabarit_target_mm=round(l_raw, 4),
        l_source_label="NC1",
        l_assembled_mm=round(l_raw, 4),
        h_target_mm=round(dstv.header.height, 4),
        h_assembled_mm=round(dstv.header.height, 4),
        w_target_mm=round(dstv.header.flange_width, 4),
        w_assembled_mm=round(dstv.header.flange_width, 4),
        max_gabarit_error_mm=0.0,
        flange_web_contact_error_mm=0.0,
        passed=True,
    )
    return plates, ver


# ─── 100% Geometric Plate Deduplication & Excel T/C/L Quantity Reporting ─────

_ROLE_LABELS_RO = {
    "talpa": "Talpă",
    "capac": "Capac",
    "flange_top": "Capac",
    "flange_bot": "Capac",
    "web": "Laterală",
    "single_plate": "Placă",
    "bom_plate": "Placă BOM",
}


def _point_to_ring_edge_dist(px: float, py: float, ring: Ring) -> float:
    n = len(ring)
    if n == 0:
        return 1e9
    best = 1e9
    for i in range(n):
        ax, ay = ring[i]
        bx, by = ring[(i + 1) % n]
        d = _dist_point_to_segment(px, py, ax, ay, bx, by)
        if d < best:
            best = d
    return best


def _ring_bidirectional_distance(r1: Ring, r2: Ring) -> float:
    if not r1 or not r2:
        return 1e9
    d12 = max(_point_to_ring_edge_dist(x, y, r2) for x, y in r1)
    d21 = max(_point_to_ring_edge_dist(x, y, r1) for x, y in r2)
    return max(d12, d21)


def _transform_point_sym(
    x: float,
    y: float,
    L: float,
    W: float,
    sym_op: str,
) -> tuple[float, float]:
    if sym_op == "identity":
        return (x, y)
    if sym_op == "rot180":
        return (L - x, W - y)
    if sym_op == "flip_x":
        return (L - x, y)
    if sym_op == "flip_y":
        return (x, W - y)
    if sym_op == "rot90":
        return (W - y, x)
    if sym_op == "rot270":
        return (y, L - x)
    return (x, y)


def _transform_angle_sym(ang_deg: float, sym_op: str) -> float:
    a = ang_deg % 180.0
    if sym_op in ("identity", "rot180"):
        return a
    if sym_op in ("flip_x", "flip_y"):
        return (-a) % 180.0
    if sym_op == "rot90":
        return (a + 90.0) % 180.0
    if sym_op == "rot270":
        return (a - 90.0) % 180.0
    return a


def _transform_ring_sym(ring: Ring, L: float, W: float, sym_op: str) -> Ring:
    return [_transform_point_sym(x, y, L, W, sym_op) for x, y in ring]


def _hole_center_and_dims(hole: DSTVHole) -> tuple[str, float, float, float, float, float]:
    kind, cx, cy, w, l, ang = _expected_hole_geom(hole)
    return (kind, cx, cy, w, l, ang % 180.0)


def _match_rings_set_sym(
    refs: list[Ring],
    cands: list[Ring],
    L: float,
    W: float,
    sym_op: str,
    tol: float,
) -> bool:
    if len(refs) != len(cands):
        return False
    if not refs:
        return True
    used: set[int] = set()
    for rr in refs:
        tr = _transform_ring_sym(rr, L, W, sym_op)
        tr_area = abs(polygon_area(tr))
        best_i = None
        for idx, cr in enumerate(cands):
            if idx in used:
                continue
            if abs(tr_area - abs(polygon_area(cr))) > 0.5:
                continue
            if _ring_bidirectional_distance(tr, cr) <= tol:
                best_i = idx
                break
        if best_i is None:
            return False
        used.add(best_i)
    return True


def _match_holes_set_sym(
    ref_holes: list[DSTVHole],
    cand_holes: list[DSTVHole],
    L: float,
    W: float,
    sym_op: str,
    tol: float,
) -> bool:
    if len(ref_holes) != len(cand_holes):
        return False
    if not ref_holes:
        return True
    cand_tuples = [_hole_center_and_dims(h) for h in cand_holes]
    used: set[int] = set()
    for rh in ref_holes:
        rk, rcx, rcy, rw, rl, rang = _hole_center_and_dims(rh)
        tcx, tcy = _transform_point_sym(rcx, rcy, L, W, sym_op)
        tang = _transform_angle_sym(rang, sym_op)
        best_i = None
        for idx, (ck, ccx, ccy, cw, cl, cang) in enumerate(cand_tuples):
            if idx in used or ck != rk:
                continue
            if abs(rw - cw) > tol or abs(rl - cl) > tol:
                continue
            if math.hypot(tcx - ccx, tcy - ccy) > tol:
                continue
            if rk != "circle":
                ang_diff = abs((tang - cang + 90.0) % 180.0 - 90.0)
                if ang_diff > 0.5:
                    continue
            best_i = idx
            break
        if best_i is None:
            return False
        used.add(best_i)
    return True


def _match_plates_geometry(
    p_ref: ExtractedPlate,
    p_cand: ExtractedPlate,
    tol: float = 0.02,
) -> Optional[str]:
    """Return matching symmetry transform op if p_ref -> p_cand is 100% geometrically identical."""
    ref_role = "capac" if p_ref.role in ("flange_top", "flange_bot") else p_ref.role
    cand_role = "capac" if p_cand.role in ("flange_top", "flange_bot") else p_cand.role
    if ref_role != cand_role:
        return None
    if abs(p_ref.thickness - p_cand.thickness) > 0.01:
        return None
    if (p_ref.material or "").strip().upper() != (p_cand.material or "").strip().upper():
        return None
    if len(p_ref.holes) != len(p_cand.holes) or len(p_ref.inner_contours) != len(p_cand.inner_contours):
        return None
    if abs(p_ref.area - p_cand.area) > 0.5:
        return None

    r0, r1, r2, r3 = p_ref.bounds
    c0, c1, c2, c3 = p_cand.bounds
    Lr, Wr = r2 - r0, r3 - r1
    Lc, Wc = c2 - c0, c3 - c1

    ops: list[str] = []
    if abs(Lr - Lc) <= 0.01 and abs(Wr - Wc) <= 0.01:
        ops.extend(["identity", "rot180", "flip_x", "flip_y"])
    if abs(Lr - Wc) <= 0.01 and abs(Wr - Lc) <= 0.01:
        ops.extend(["rot90", "rot270"])
    if not ops:
        return None

    for op in ops:
        t_ext = _transform_ring_sym(p_ref.exterior, Lr, Wr, op)
        if _ring_bidirectional_distance(t_ext, p_cand.exterior) > tol:
            continue
        if not _match_rings_set_sym(p_ref.inner_contours, p_cand.inner_contours, Lr, Wr, op, tol):
            continue
        if not _match_holes_set_sym(p_ref.holes, p_cand.holes, Lr, Wr, op, tol):
            continue
        return op
    return None


def _clone_plate_as_canonical(
    plate: ExtractedPlate,
    new_name: str,
    role_code: str,
    qty_per_cheson: int,
    merged_sources: list[str],
) -> ExtractedPlate:
    return ExtractedPlate(
        name=new_name,
        folder=plate.folder,
        role=plate.role,
        face=plate.face,
        thickness=plate.thickness,
        piece_id=plate.piece_id,
        piece_clean=plate.piece_clean,
        material=plate.material,
        quantity=plate.quantity,
        exterior=list(plate.exterior),
        holes=list(plate.holes),
        inner_contours=[list(ir) for ir in plate.inner_contours],
        markings=list(plate.markings),
        source_bounds=plate.source_bounds,
        removed_residuals=plate.removed_residuals,
        qty_per_cheson=qty_per_cheson,
        role_code=role_code,
        merged_sources=merged_sources,
    )


def deduplicate_cheson_plates(
    raw_plates: list[ExtractedPlate],
    cfg: ChesonConfig,
) -> tuple[list[ExtractedPlate], list[tuple[int, str]]]:
    """Group geometrically identical plates per role (T/C/L) and normalize names.

    Returns `(unique_plates, instance_mapping)` where `instance_mapping[raw_idx]`
    is `(unique_idx, sym_op_from_canonical_to_raw)`.
    """
    if not raw_plates:
        return [], []

    def _norm_role(role: str) -> str:
        return "capac" if role in ("flange_top", "flange_bot") else role

    role_order = ("talpa", "capac", "web", "single_plate")
    role_indices: dict[str, list[int]] = {r: [] for r in role_order}
    other_indices: list[int] = []
    for idx, p in enumerate(raw_plates):
        nr = _norm_role(p.role)
        if nr in role_indices:
            role_indices[nr].append(idx)
        else:
            other_indices.append(idx)

    letter_map = {
        "talpa": "T",
        "capac": (cfg.flange_letter or "C").upper(),
        "web": (cfg.web_letter or "L").upper(),
        "single_plate": "P",
    }

    unique_plates: list[ExtractedPlate] = []
    mapping: list[Optional[tuple[int, str]]] = [None] * len(raw_plates)

    for nr in role_order:
        idxs = role_indices[nr]
        if not idxs:
            continue
        letter = letter_map[nr]
        clusters: list[tuple[int, list[tuple[int, str]]]] = []
        for raw_idx in idxs:
            if not cfg.deduplicate_identical:
                clusters.append((raw_idx, [(raw_idx, "identity")]))
                continue
            matched_cluster = False
            for c_idx, (rep_idx, members) in enumerate(clusters):
                op = _match_plates_geometry(raw_plates[rep_idx], raw_plates[raw_idx])
                if op is not None:
                    members.append((raw_idx, op))
                    matched_cluster = True
                    break
            if not matched_cluster:
                clusters.append((raw_idx, [(raw_idx, "identity")]))

        k_classes = len(clusters)
        for cls_num, (rep_idx, members) in enumerate(clusters, start=1):
            rep = raw_plates[rep_idx]
            if nr == "single_plate":
                role_code = letter if k_classes == 1 else f"{letter}{cls_num}"
                new_name = rep.piece_clean if k_classes == 1 else f"{rep.piece_clean}_{cls_num}"
            else:
                role_code = letter if k_classes == 1 else f"{letter}{cls_num}"
                new_name = f"{rep.piece_clean}{role_code}"

            merged_src = [f"{raw_plates[m_idx].name}@{raw_plates[m_idx].face}" for m_idx, _ in members]
            u_idx = len(unique_plates)
            unique_plates.append(
                _clone_plate_as_canonical(
                    rep,
                    new_name=new_name,
                    role_code=role_code,
                    qty_per_cheson=len(members),
                    merged_sources=merged_src,
                )
            )
            for m_idx, op in members:
                mapping[m_idx] = (u_idx, op)

    for raw_idx in other_indices:
        p = raw_plates[raw_idx]
        u_idx = len(unique_plates)
        unique_plates.append(
            _clone_plate_as_canonical(
                p,
                new_name=p.name,
                role_code=p.role_code or "P",
                qty_per_cheson=1,
                merged_sources=[f"{p.name}@{p.face}"],
            )
        )
        mapping[raw_idx] = (u_idx, "identity")

    final_mapping: list[tuple[int, str]] = [
        m if m is not None else (0, "identity") for m in mapping
    ]
    return unique_plates, final_mapping


def _expand_audits_for_3d_sim(
    raw_plates: list[ExtractedPlate],
    unique_plates: list[ExtractedPlate],
    instance_mapping: list[tuple[int, str]],
    unique_audits: list[DrawnDxfAudit],
    unique_re_read_rings: dict[str, Ring],
) -> tuple[list[DrawnDxfAudit], dict[str, Ring]]:
    """Expand deduplicated DXF audits/rings back onto every physical plate instance for 3D simulation."""
    expanded_audits: list[DrawnDxfAudit] = []
    expanded_rings: dict[str, Ring] = {}

    for raw_idx, raw_p in enumerate(raw_plates):
        u_idx, sym_op = instance_mapping[raw_idx]
        canon = unique_plates[u_idx]
        c_aud = unique_audits[u_idx]
        c_ring = unique_re_read_rings.get(canon.name, canon.exterior)

        c0, c1, c2, c3 = canon.bounds
        L_c, W_c = c2 - c0, c3 - c1
        inst_ring = _transform_ring_sym(c_ring, L_c, W_c, sym_op)
        expanded_rings[raw_p.name] = inst_ring

        inst_holes: list[DrawnHoleInfo] = []
        for dh in c_aud.holes_drawn:
            tcx, tcy = _transform_point_sym(dh.cx, dh.cy, L_c, W_c, sym_op)
            tang = _transform_angle_sym(dh.angle_deg, sym_op)
            inst_holes.append(
                DrawnHoleInfo(
                    kind=dh.kind,
                    cx=tcx,
                    cy=tcy,
                    width_mm=dh.width_mm,
                    length_mm=dh.length_mm,
                    angle_deg=tang,
                    area_mm2=dh.area_mm2,
                )
            )

        ib0, ib1, ib2, ib3 = polygon_bounds(inst_ring) if inst_ring else (0.0, 0.0, c_aud.length_mm, c_aud.width_mm)
        expanded_audits.append(
            DrawnDxfAudit(
                plate_name=raw_p.name,
                role=raw_p.role,
                face=raw_p.face,
                thickness_mm=raw_p.thickness,
                dxf_errors=c_aud.dxf_errors,
                entity_layers=list(c_aud.entity_layers),
                extra_layers_created=list(c_aud.extra_layers_created),
                layer_zero_ok=c_aud.layer_zero_ok,
                outer_closed=c_aud.outer_closed,
                origin_err_mm=c_aud.origin_err_mm,
                length_mm=round(ib2 - ib0, 4),
                width_mm=round(ib3 - ib1, 4),
                outer_area_mm2=c_aud.outer_area_mm2,
                net_area_mm2=c_aud.net_area_mm2,
                vertex_count=c_aud.vertex_count,
                holes_drawn=inst_holes,
                inners_count=c_aud.inners_count,
                dim_err_mm=c_aud.dim_err_mm,
                hole_pos_err_mm=c_aud.hole_pos_err_mm,
                hole_dim_err_mm=c_aud.hole_dim_err_mm,
                holes_matched=c_aud.holes_matched,
                holes_expected=c_aud.holes_expected,
                passed=c_aud.passed,
                issues=list(c_aud.issues),
            )
        )

    return expanded_audits, expanded_rings


def find_cheson_report_xlsx(
    search_root: Path,
    explicit_path: Optional[str | Path] = None,
) -> Optional[Path]:
    """Locate raport_CHESON.xlsx either via explicit path or inside search_root / parent."""
    if explicit_path:
        p = Path(str(explicit_path).strip())
        if p.is_file():
            return p.resolve()

    candidates: list[Path] = []
    dirs = [search_root]
    if search_root.parent != search_root:
        dirs.append(search_root.parent)
    for d in dirs:
        if not d.is_dir():
            continue
        for f in sorted(d.iterdir()):
            if not f.is_file() or f.suffix.lower() not in (".xlsx", ".xlsm"):
                continue
            nl = f.name.lower()
            if nl == "raport_cheson.xlsx":
                return f.resolve()
            if "raport" in nl and "cheson" in nl and not nl.startswith("raport_piese"):
                candidates.append(f.resolve())
    return candidates[0] if candidates else None


def load_cheson_quantities_from_excel(xlsx_path: Optional[Path]) -> dict[str, int]:
    """Parse cheson piece quantities across all sheets of raport_CHESON.xlsx."""
    if xlsx_path is None or not Path(xlsx_path).is_file() or openpyxl is None:
        return {}

    qty_map: dict[str, int] = {}
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            piece_col: Optional[int] = None
            qty_col: Optional[int] = None
            header_row_idx = 0

            rows_iter = list(ws.iter_rows(values_only=True))
            for r_idx, row in enumerate(rows_iter[:15]):
                if not row:
                    continue
                texts = [str(c).strip().lower() if c is not None else "" for c in row]
                p_c = None
                q_c = None
                for c_idx, txt in enumerate(texts):
                    if not txt:
                        continue
                    if p_c is None and any(k in txt for k in ("pies", "pos", "denumire", "cod", "cheson", "part", "name", "reper")):
                        p_c = c_idx
                    elif q_c is None and any(k in txt for k in ("cant", "buc", "qty", "nr", "numar")):
                        q_c = c_idx
                if p_c is not None and q_c is not None:
                    piece_col, qty_col, header_row_idx = p_c, q_c, r_idx + 1
                    break

            if piece_col is None or qty_col is None:
                piece_col, qty_col, header_row_idx = 0, 1, 0

            for row in rows_iter[header_row_idx:]:
                if not row or len(row) <= max(piece_col, qty_col):
                    continue
                raw_name = row[piece_col]
                raw_qty = row[qty_col]
                if raw_name is None or raw_qty is None:
                    continue
                name_s = str(raw_name).strip()
                if not name_s or any(k in name_s.lower() for k in ("total", "denumire", "sumar")):
                    continue
                if isinstance(raw_qty, (int, float)):
                    q_val = int(round(float(raw_qty)))
                else:
                    q_parsed = _parse_float_opt(str(raw_qty))
                    q_val = int(round(q_parsed)) if q_parsed is not None else 0
                if q_val <= 0:
                    continue
                key = clean_piece_name(name_s, lowercase=True)
                if key:
                    qty_map[key] = qty_map.get(key, 0) + q_val
    finally:
        wb.close()
    return qty_map


def format_cheson_tcl_summary(plate_rows: list[dict[str, Any]]) -> dict[str, str]:
    """Build grouped T / C / L quantity summary strings for a single cheson."""
    t_parts: list[str] = []
    c_parts: list[str] = []
    l_parts: list[str] = []

    for row in plate_rows:
        rc = (row.get("role_code") or "").strip().upper()
        tot = int(row.get("total_qty") or 0)
        if not rc:
            continue
        item_str = f"{rc}: {tot}"
        if rc.startswith("T"):
            t_parts.append(item_str)
        elif rc.startswith("C"):
            c_parts.append(item_str)
        elif rc.startswith("L"):
            l_parts.append(item_str)

    all_parts = t_parts + c_parts + l_parts
    return {
        "talpa_str": ", ".join(t_parts) if t_parts else "-",
        "capac_str": ", ".join(c_parts) if c_parts else "-",
        "laterala_str": ", ".join(l_parts) if l_parts else "-",
        "full_summary": " | ".join(all_parts) if all_parts else "-",
    }


def write_cheson_quantity_report_xlsx(
    output_path: Path,
    cheson_summaries: list[dict[str, Any]],
    plate_entries: list[dict[str, Any]],
) -> Optional[Path]:
    """Write formatted Raport_Piese_Cheson.xlsx (Sumar Chesoane T-C-L, Lista Piese Debitare, Pe Grosimi)."""
    if openpyxl is None:
        return None

    wb = openpyxl.Workbook()
    header_fill = PatternFill("solid", fgColor="1E293B")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill("solid", fgColor="E2E8F0")
    total_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    pass_fill = PatternFill("solid", fgColor="DCFCE7")
    check_fill = PatternFill("solid", fgColor="FEF3C7")
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def _style_sheet(ws, headers: list[str], rows_data: list[list[Any]], center_cols: set[int] = frozenset()) -> None:
        ws.append(headers)
        for col_i in range(1, len(headers) + 1):
            cell = ws.cell(1, col_i)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[1].height = 24

        for r_idx, r_vals in enumerate(rows_data, start=2):
            ws.append(r_vals)
            for col_i in range(1, len(headers) + 1):
                c = ws.cell(r_idx, col_i)
                c.border = border
                c.alignment = Alignment(
                    horizontal=("center" if col_i in center_cols else "left"),
                    vertical="center",
                )

        ws.freeze_panes = "A2"
        for col_i in range(1, len(headers) + 1):
            col_letter = openpyxl.utils.get_column_letter(col_i)
            max_len = max(
                len(str(ws.cell(r, col_i).value or ""))
                for r in range(1, ws.max_row + 1)
            )
            ws.column_dimensions[col_letter].width = min(48, max(11, max_len + 3))

    # Sheet 1: Sumar Chesoane (T-C-L)
    ws1 = wb.active
    ws1.title = "Sumar Chesoane (T-C-L)"
    s1_headers = [
        "Cheson", "Folder", "Profil", "Tip Cheson", "Nr. Chesoane", "Sursă Cant.",
        "Talpă (T)", "Capace (C)", "Laterale (L)", "Sumar Complet (T / C / L)",
        "Piese Unice", "Total Bucăți", "Suprapunere PDF (%)", "Status Audit 3D",
    ]
    s1_rows = [
        [
            cs.get("cheson_id", ""),
            cs.get("folder", ""),
            cs.get("profile", ""),
            cs.get("detected_type_label", ""),
            cs.get("cheson_qty", 1),
            cs.get("qty_source", ""),
            cs.get("talpa_str", "-"),
            cs.get("capac_str", "-"),
            cs.get("laterala_str", "-"),
            cs.get("full_summary", "-"),
            cs.get("unique_plates_count", 0),
            cs.get("total_pieces_count", 0),
            (f"{cs['overlay_pct']:.2f}%" if cs.get("overlay_pct") is not None else "-"),
            cs.get("status", ""),
        ]
        for cs in cheson_summaries
    ]
    _style_sheet(ws1, s1_headers, s1_rows, center_cols={5, 11, 12, 13, 14})
    for r_i in range(2, ws1.max_row + 1):
        st_cell = ws1.cell(r_i, 14)
        st_val = str(st_cell.value or "").upper()
        if st_val == "PASS":
            st_cell.fill = pass_fill
        elif st_val:
            st_cell.fill = check_fill

    if s1_rows:
        tot_r = ws1.max_row + 1
        tot_vals = [
            "TOTAL",
            "",
            "",
            f"{len(cheson_summaries)} chesoane distincte",
            sum(int(cs.get("cheson_qty", 0)) for cs in cheson_summaries),
            "",
            "",
            "",
            "",
            "",
            sum(int(cs.get("unique_plates_count", 0)) for cs in cheson_summaries),
            sum(int(cs.get("total_pieces_count", 0)) for cs in cheson_summaries),
            "",
            "",
        ]
        ws1.append(tot_vals)
        for col_i in range(1, len(s1_headers) + 1):
            c = ws1.cell(tot_r, col_i)
            c.fill = total_fill
            c.font = total_font
            c.border = border
            if col_i in {5, 11, 12, 13}:
                c.alignment = Alignment(horizontal="center", vertical="center")

    # Sheet 2: Lista Piese Debitare
    ws2 = wb.create_sheet("Lista Piese Debitare")
    s2_headers = [
        "Folder Cheson", "Cheson", "Denumire Piesă", "Cod Rol", "Tip Piesă",
        "Grosime (mm)", "Material", "Lungime (mm)", "Lățime (mm)",
        "Buc/Cheson", "Nr. Chesoane", "Total Bucăți", "Sursă Cant.",
        "Fișier CAD", "Piese Identice Înglobate",
    ]
    s2_rows = [
        [
            e.get("cheson_folder", ""),
            e.get("pos_no", ""),
            e.get("plate_name", ""),
            e.get("role_code", ""),
            e.get("role_label", ""),
            e.get("thickness_mm"),
            e.get("material", ""),
            e.get("length_mm"),
            e.get("width_mm"),
            e.get("qty_per_cheson", e.get("qty_per_assembly", 1)),
            e.get("assembly_qty", 1),
            e.get("total_qty", 1),
            e.get("qty_source", ""),
            e.get("output_file", ""),
            e.get("notes", ""),
        ]
        for e in plate_entries
    ]
    _style_sheet(ws2, s2_headers, s2_rows, center_cols={4, 6, 8, 9, 10, 11, 12})

    # Sheet 3: Pe Grosimi
    ws3 = wb.create_sheet("Pe Grosimi")
    grouped: dict[tuple[float, str], dict[str, Any]] = {}
    for e in plate_entries:
        t_v = float(e.get("thickness_mm") or 0.0)
        mat = (e.get("material") or "").strip()
        key = (t_v, mat)
        if key not in grouped:
            grouped[key] = {"t": t_v, "mat": mat, "unique_cnt": 0, "total_pcs": 0}
        grouped[key]["unique_cnt"] += 1
        grouped[key]["total_pcs"] += int(e.get("total_qty") or 0)

    s3_headers = ["Grosime (mm)", "Folder Grosime", "Material", "Nr. Piese Unice", "Total Bucăți"]
    s3_rows = [
        [g["t"], f"T{int(round(g['t']))}" if g["t"] > 0 else "-", g["mat"], g["unique_cnt"], g["total_pcs"]]
        for _, g in sorted(grouped.items(), key=lambda kv: (kv[0][0], kv[0][1]))
    ]
    _style_sheet(ws3, s3_headers, s3_rows, center_cols={1, 2, 4, 5})

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    return output_path


# ─── CAD DXF R2018 & ODA DWG Export ──────────────────────────────────────────

def _interior_label_point(ring: Ring) -> tuple[float, float]:
    minx, miny, maxx, maxy = polygon_bounds(ring)
    cx, cy = 0.5 * (minx + maxx), 0.5 * (miny + maxy)
    if point_in_polygon(cx, cy, ring, tol=0.0):
        return (cx, cy)
    for frac in (0.25, 0.75, 0.15, 0.85):
        px = minx + frac * (maxx - minx)
        if point_in_polygon(px, cy, ring, tol=0.0):
            return (px, cy)
    return (cx, cy)


def _add_slot_lwpolyline(msp, hole: DSTVHole, layer: str) -> None:
    """Emit closed CCW LWPOLYLINE for DSTV BO 'l' slotted/rounded-rect holes."""
    r = 0.5 * hole.diameter
    if r <= 1e-6:
        return
    dx_s = float(hole.slot_length)
    dy_s = float(hole.slot_width)
    rad = math.radians(hole.slot_angle)
    ux, uy = math.cos(rad), math.sin(rad)
    vx, vy = -uy, ux

    # Case B: 2D rounded-rectangle slot (both dx_s > 0.05 and dy_s > 0.05)
    if abs(dx_s) > 0.05 and abs(dy_s) > 0.05:
        k0 = (hole.x, hole.y)
        k1 = (k0[0] + dx_s * ux, k0[1] + dx_s * uy)
        k2 = (k1[0] + dy_s * vx, k1[1] + dy_s * vy)
        k3 = (k0[0] + dy_s * vx, k0[1] + dy_s * vy)
        bulge90 = math.tan(math.pi / 8.0)
        pts = [
            (k0[0] - r * vx, k0[1] - r * vy, 0.0),
            (k1[0] - r * vx, k1[1] - r * vy, bulge90),
            (k1[0] + r * ux, k1[1] + r * uy, 0.0),
            (k2[0] + r * ux, k2[1] + r * uy, bulge90),
            (k2[0] + r * vx, k2[1] + r * vy, 0.0),
            (k3[0] + r * vx, k3[1] + r * vy, bulge90),
            (k3[0] - r * ux, k3[1] - r * uy, 0.0),
            (k0[0] - r * ux, k0[1] - r * uy, bulge90),
        ]
        msp.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": layer})
        return

    # Case A: 1D oval / stadium slot (one extension > 0.05)
    c1 = (hole.x, hole.y)
    c2 = (hole.x + dx_s * ux + dy_s * vx, hole.y + dx_s * uy + dy_s * vy)
    span = math.hypot(c2[0] - c1[0], c2[1] - c1[1])
    if span < 0.05:
        msp.add_circle(c1, radius=r, dxfattribs={"layer": layer})
        return

    ax = (c2[0] - c1[0]) / span
    ay = (c2[1] - c1[1]) / span
    nx, ny = -ay, ax
    pts = [
        (c1[0] - r * nx, c1[1] - r * ny, 0.0),
        (c2[0] - r * nx, c2[1] - r * ny, 1.0),
        (c2[0] + r * nx, c2[1] + r * ny, 0.0),
        (c1[0] + r * nx, c1[1] + r * ny, 1.0),
    ]
    msp.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": layer})


def write_plate_dxf(
    plate: ExtractedPlate,
    output_path: str | Path,
    use_nesting_layers: bool = False,
    include_marking: bool = False,
    marking_height: float = 20.0,
) -> Path:
    if ezdxf is None:
        raise RuntimeError("Biblioteca 'ezdxf' nu este disponibilă.")

    doc = ezdxf.new("R2018", setup=True)
    doc.header["$INSUNITS"] = 4
    doc.header["$MEASUREMENT"] = 1
    doc.header["$DIMLFAC"] = 1.0

    if use_nesting_layers:
        for lname, col in (("CUT_OUTER", 7), ("CUT_INNER", 4), ("MARKING", 2)):
            if lname in doc.layers:
                doc.layers.get(lname).color = col
            else:
                doc.layers.add(lname, color=col)
        outer_layer, inner_layer, mark_layer = "CUT_OUTER", "CUT_INNER", "MARKING"
    else:
        outer_layer = inner_layer = mark_layer = "0"

    msp = doc.modelspace()
    ext_pts = clean_ring(plate.exterior, min_seg_len=0.05, ccw=True)
    if len(ext_pts) >= 3:
        msp.add_lwpolyline(ext_pts, format="xy", close=True, dxfattribs={"layer": outer_layer})

    for ik_ring in plate.inner_contours:
        ik_pts = clean_ring(ik_ring, min_seg_len=0.05, ccw=False)
        if len(ik_pts) >= 3:
            msp.add_lwpolyline(ik_pts, format="xy", close=True, dxfattribs={"layer": inner_layer})

    for hole in plate.holes:
        if hole.hole_type == "l" and (abs(hole.slot_length) > 0.05 or abs(hole.slot_width) > 0.05):
            _add_slot_lwpolyline(msp, hole, layer=inner_layer)
        elif hole.diameter > 0.0:
            msp.add_circle((hole.x, hole.y), radius=0.5 * hole.diameter, dxfattribs={"layer": inner_layer})

    if include_marking:
        _, miny, _, maxy = plate.bounds
        text_h = min(float(marking_height), max(4.0, (maxy - miny) * 0.25))
        lx, ly = _interior_label_point(plate.exterior)
        label_str = " ".join(
            p for p in [plate.piece_id or plate.name.upper(), plate.name, plate.material, f"T{int(round(plate.thickness))}"] if p
        )
        txt = msp.add_text(label_str, height=text_h, dxfattribs={"layer": mark_layer})
        txt.set_placement((lx, ly), align=TextEntityAlignment.MIDDLE_CENTER)
        for mark in plate.markings:
            if mark.text:
                mt = msp.add_text(
                    mark.text,
                    height=max(4.0, mark.height or text_h),
                    dxfattribs={"layer": mark_layer, "rotation": mark.angle},
                )
                mt.set_placement((mark.x, mark.y), align=TextEntityAlignment.LEFT)

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.saveas(str(out_path))
    return out_path


# ─── Post-Draw DXF Re-Verification & 3D Virtual Assembly Simulation ───────────

def _lwpolyline_geometric_info(
    lwpoly,
) -> tuple[Ring, tuple[float, float, float, float], float, Optional[DrawnHoleInfo]]:
    """Compute exact vertices, bulge-aware bbox & area, and classify hole LWPOLYLINEs."""
    raw = [(float(p[0]), float(p[1]), float(p[2])) for p in lwpoly.get_points("xyb")]
    n = len(raw)
    ring: Ring = [(p[0], p[1]) for p in raw]
    if n < 2:
        return ring, (0.0, 0.0, 0.0, 0.0), 0.0, None

    xs = [p[0] for p in raw]
    ys = [p[1] for p in raw]
    shoelace = 0.0
    seg_area = 0.0
    bulges: list[tuple[int, float]] = []

    for i in range(n):
        x1, y1, b = raw[i]
        x2, y2, _ = raw[(i + 1) % n]
        shoelace += x1 * y2 - x2 * y1
        if abs(b) > 1e-6:
            bulges.append((i, b))
            dx, dy = x2 - x1, y2 - y1
            chord = math.hypot(dx, dy)
            if chord > 1e-8:
                theta = 4.0 * math.atan(b)
                tan_half = math.tan(0.5 * theta)
                sin_half = math.sin(0.5 * theta)
                r = abs(chord / (2.0 * sin_half)) if abs(sin_half) > 1e-9 else 0.0
                if r > 0.0:
                    seg_area += 0.5 * (r * r) * (theta - math.sin(theta))
                    d_center = (chord / (2.0 * tan_half)) if abs(tan_half) > 1e-9 else 0.0
                    nx, ny = -dy / chord, dx / chord
                    cx = 0.5 * (x1 + x2) + d_center * nx
                    cy = 0.5 * (y1 + y2) + d_center * ny
                    a1 = math.atan2(y1 - cy, x1 - cx)
                    for c_ang in (0.0, 0.5 * math.pi, math.pi, 1.5 * math.pi):
                        if theta > 0.0:
                            rel = (c_ang - a1) % (2.0 * math.pi)
                            inside_arc = 1e-7 < rel < theta - 1e-7
                        else:
                            rel = (a1 - c_ang) % (2.0 * math.pi)
                            inside_arc = 1e-7 < rel < (-theta) - 1e-7
                        if inside_arc:
                            xs.append(cx + r * math.cos(c_ang))
                            ys.append(cy + r * math.sin(c_ang))

    # Signed area: shoelace + 2*seg_area (both multiplied by 0.5)
    area = abs(0.5 * shoelace + seg_area)
    bbox = (min(xs), min(ys), max(xs), max(ys))
    cx_box = 0.5 * (bbox[0] + bbox[2])
    cy_box = 0.5 * (bbox[1] + bbox[3])

    hole_info: Optional[DrawnHoleInfo] = None
    if lwpoly.closed and n == 4 and len(bulges) == 2 and all(abs(abs(b) - 1.0) < 0.02 for _, b in bulges):
        # 1D oval slot: 2 semicircular end arcs (bulge +-1.0) + 2 parallel straight sides
        b_idx0 = bulges[0][0]
        p_a = raw[b_idx0]
        p_b = raw[(b_idx0 + 1) % 4]
        width = math.hypot(p_b[0] - p_a[0], p_b[1] - p_a[1])
        straight_segs = [
            (raw[i], raw[(i + 1) % 4])
            for i in range(4)
            if abs(raw[i][2]) < 1e-4
        ]
        if straight_segs:
            s0, s1 = straight_segs[0]
            sdx, sdy = s1[0] - s0[0], s1[1] - s0[1]
            span = math.hypot(sdx, sdy)
            ang = math.degrees(math.atan2(sdy, sdx)) % 180.0
            hole_info = DrawnHoleInfo(
                kind="slot",
                cx=round(cx_box, 4),
                cy=round(cy_box, 4),
                width_mm=round(width, 4),
                length_mm=round(width + span, 4),
                angle_deg=round(ang, 3),
                area_mm2=round(area, 4),
            )
    elif lwpoly.closed and n == 8 and len(bulges) == 4 and all(abs(abs(b) - (math.sqrt(2.0) - 1.0)) < 0.02 for _, b in bulges):
        # 2D rounded-rectangle slot: 4 90-deg corner arcs + 4 straight sides
        arc_chords = [
            math.hypot(raw[(i + 1) % 8][0] - raw[i][0], raw[(i + 1) % 8][1] - raw[i][1])
            for i, _ in bulges
        ]
        diam = math.sqrt(2.0) * (sum(arc_chords) / len(arc_chords))
        straight_lens = [
            math.hypot(raw[(i + 1) % 8][0] - raw[i][0], raw[(i + 1) % 8][1] - raw[i][1])
            for i in range(8)
            if abs(raw[i][2]) < 1e-4
        ]
        if len(straight_lens) >= 2:
            s_pair = sorted(straight_lens)
            w_s = diam + 0.5 * (s_pair[0] + s_pair[1])
            l_s = diam + 0.5 * (s_pair[-2] + s_pair[-1])
            hole_info = DrawnHoleInfo(
                kind="rect_slot",
                cx=round(cx_box, 4),
                cy=round(cy_box, 4),
                width_mm=round(w_s, 4),
                length_mm=round(l_s, 4),
                angle_deg=0.0,
                area_mm2=round(area, 4),
            )

    return ring, bbox, area, hole_info


def _expected_hole_geom(hole: DSTVHole) -> tuple[str, float, float, float, float, float]:
    """Return (kind, exp_cx, exp_cy, exp_width, exp_length, exp_angle_mod180)."""
    d = float(hole.diameter)
    if hole.hole_type == "l" and (abs(hole.slot_length) > 0.05 or abs(hole.slot_width) > 0.05):
        rad = math.radians(hole.slot_angle)
        ux, uy = math.cos(rad), math.sin(rad)
        vx, vy = -uy, ux
        dx_s, dy_s = float(hole.slot_length), float(hole.slot_width)
        exp_cx = hole.x + 0.5 * (dx_s * ux + dy_s * vx)
        exp_cy = hole.y + 0.5 * (dx_s * uy + dy_s * vy)
        if abs(dx_s) > 0.05 and abs(dy_s) > 0.05:
            return ("rect_slot", exp_cx, exp_cy, d + min(abs(dx_s), abs(dy_s)), d + max(abs(dx_s), abs(dy_s)), hole.slot_angle % 180.0)
        span = math.hypot(dx_s * ux + dy_s * vx, dx_s * uy + dy_s * vy)
        ang = math.degrees(math.atan2(dx_s * uy + dy_s * vy, dx_s * ux + dy_s * vx)) % 180.0
        return ("slot", exp_cx, exp_cy, d, d + span, ang)
    return ("circle", hole.x, hole.y, d, d, 0.0)


def inspect_drawn_dxf(
    dxf_path: Path,
    plate: ExtractedPlate,
    expect_layer_zero: bool = True,
) -> tuple[DrawnDxfAudit, Ring]:
    """Re-read a written DXF file via ezdxf and audit layers, dimensions, origin, contours, and holes."""
    if ezdxf is None:
        return (
            DrawnDxfAudit(
                plate_name=plate.name,
                role=plate.role,
                face=plate.face,
                thickness_mm=plate.thickness,
                dxf_errors=1,
                entity_layers=[],
                extra_layers_created=[],
                layer_zero_ok=False,
                outer_closed=False,
                origin_err_mm=999.0,
                length_mm=0.0,
                width_mm=0.0,
                outer_area_mm2=0.0,
                net_area_mm2=0.0,
                vertex_count=0,
                holes_drawn=[],
                inners_count=0,
                dim_err_mm=999.0,
                hole_pos_err_mm=999.0,
                hole_dim_err_mm=999.0,
                holes_matched=0,
                holes_expected=len(plate.holes),
                passed=False,
                issues=["ezdxf indisponibil"],
            ),
            list(plate.exterior),
        )

    doc = ezdxf.readfile(str(dxf_path))
    aud = doc.audit()
    dxf_errors = len(aud.errors)
    extra_layers = sorted(
        l.dxf.name for l in doc.layers if l.dxf.name not in ("0", "Defpoints")
    )
    msp = doc.modelspace()
    entity_layers = sorted({e.dxf.layer for e in msp})
    layer_zero_ok = (not expect_layer_zero) or (
        set(entity_layers).issubset({"0"}) and len(extra_layers) == 0
    )

    polys_data = []
    holes_drawn: list[DrawnHoleInfo] = []
    for e in msp:
        etype = e.dxftype()
        if etype == "LWPOLYLINE":
            ring, bbox, area, hinfo = _lwpolyline_geometric_info(e)
            polys_data.append((e.closed, ring, bbox, area, hinfo))
        elif etype == "CIRCLE":
            c = e.dxf.center
            d = 2.0 * float(e.dxf.radius)
            holes_drawn.append(
                DrawnHoleInfo(
                    kind="circle",
                    cx=round(float(c.x), 4),
                    cy=round(float(c.y), 4),
                    width_mm=round(d, 4),
                    length_mm=round(d, 4),
                    angle_deg=0.0,
                    area_mm2=round(math.pi * (0.5 * d) ** 2, 4),
                )
            )

    issues: list[str] = []
    if dxf_errors > 0:
        issues.append(f"Erori ezdxf audit: {dxf_errors}")
    if expect_layer_zero and not layer_zero_ok:
        issues.append(f"Layere neconforme: entități={entity_layers}, extra={extra_layers}")

    if not polys_data:
        issues.append("Lipsă contur exterior LWPOLYLINE în DXF")
        return (
            DrawnDxfAudit(
                plate_name=plate.name,
                role=plate.role,
                face=plate.face,
                thickness_mm=plate.thickness,
                dxf_errors=dxf_errors,
                entity_layers=entity_layers,
                extra_layers_created=extra_layers,
                layer_zero_ok=layer_zero_ok,
                outer_closed=False,
                origin_err_mm=999.0,
                length_mm=0.0,
                width_mm=0.0,
                outer_area_mm2=0.0,
                net_area_mm2=0.0,
                vertex_count=0,
                holes_drawn=holes_drawn,
                inners_count=0,
                dim_err_mm=999.0,
                hole_pos_err_mm=0.0,
                hole_dim_err_mm=0.0,
                holes_matched=0,
                holes_expected=len(plate.holes),
                passed=False,
                issues=issues,
            ),
            list(plate.exterior),
        )

    # Outer contour is largest LWPOLYLINE by area
    polys_data.sort(key=lambda item: item[3], reverse=True)
    outer_closed, outer_ring, outer_bbox, outer_area, _ = polys_data[0]
    if not outer_closed:
        issues.append("Conturul exterior NU este închis")

    l_drawn = outer_bbox[2] - outer_bbox[0]
    w_drawn = outer_bbox[3] - outer_bbox[1]
    origin_err = max(abs(outer_bbox[0]), abs(outer_bbox[1]))
    if origin_err > 0.01:
        issues.append(f"Origine DXF deplasată ({outer_bbox[0]:.3f}, {outer_bbox[1]:.3f})")

    exp_b = plate.bounds
    l_exp = exp_b[2] - exp_b[0]
    w_exp = exp_b[3] - exp_b[1]
    dim_err = max(abs(l_drawn - l_exp), abs(w_drawn - w_exp))
    if dim_err > 0.01:
        issues.append(f"Diferență gabarit DXF vs NC1: {dim_err:.4f} mm")

    inner_area_sum = 0.0
    inners_count = 0
    for is_closed, _, _, p_area, hinfo in polys_data[1:]:
        if hinfo is not None:
            holes_drawn.append(hinfo)
        else:
            inners_count += 1
            inner_area_sum += p_area

    for h in holes_drawn:
        inner_area_sum += h.area_mm2
        if not point_in_polygon(h.cx, h.cy, outer_ring, tol=0.05):
            issues.append(f"Gaură ({h.cx:.2f}, {h.cy:.2f}) în afara conturului exterior")

    # Match expected DSTV holes 1-to-1 against drawn holes
    used_idx: set[int] = set()
    hole_pos_err = 0.0
    hole_dim_err = 0.0
    holes_matched = 0
    for exp_h in plate.holes:
        ekind, ecx, ecy, ew, el, _ = _expected_hole_geom(exp_h)
        best_i = None
        best_dist = 1e9
        for i_d, dh in enumerate(holes_drawn):
            if i_d in used_idx:
                continue
            dist = math.hypot(dh.cx - ecx, dh.cy - ecy)
            if dist < best_dist:
                best_dist = dist
                best_i = i_d
        if best_i is None or best_dist > 1.0:
            issues.append(f"Gaură lipsă/deplasată la ({ecx:.2f}, {ecy:.2f})")
            hole_pos_err = max(hole_pos_err, best_dist if best_i is not None else 999.0)
            continue
        used_idx.add(best_i)
        dh = holes_drawn[best_i]
        if dh.kind != ekind:
            issues.append(f"Tip gaură greșit la ({ecx:.2f}, {ecy:.2f}): desenat {dh.kind}, așteptat {ekind}")
        d_err = max(abs(dh.width_mm - ew), abs(dh.length_mm - el))
        hole_pos_err = max(hole_pos_err, best_dist)
        hole_dim_err = max(hole_dim_err, d_err)
        if best_dist <= 0.05 and d_err <= 0.05 and dh.kind == ekind:
            holes_matched += 1
        else:
            issues.append(
                f"Abatere gaură ({ecx:.2f}, {ecy:.2f}): pos={best_dist:.3f} mm, dim={d_err:.3f} mm"
            )

    net_area = max(0.0, outer_area - inner_area_sum)
    passed = (
        dxf_errors == 0
        and layer_zero_ok
        and outer_closed
        and origin_err <= 0.01
        and dim_err <= 0.01
        and holes_matched == len(plate.holes)
        and len(holes_drawn) == len(plate.holes)
        and len(issues) == 0
    )

    return (
        DrawnDxfAudit(
            plate_name=plate.name,
            role=plate.role,
            face=plate.face,
            thickness_mm=plate.thickness,
            dxf_errors=dxf_errors,
            entity_layers=entity_layers,
            extra_layers_created=extra_layers,
            layer_zero_ok=layer_zero_ok,
            outer_closed=outer_closed,
            origin_err_mm=round(origin_err, 4),
            length_mm=round(l_drawn, 4),
            width_mm=round(w_drawn, 4),
            outer_area_mm2=round(outer_area, 2),
            net_area_mm2=round(net_area, 2),
            vertex_count=len(outer_ring),
            holes_drawn=holes_drawn,
            inners_count=inners_count,
            dim_err_mm=round(dim_err, 4),
            hole_pos_err_mm=round(hole_pos_err, 4),
            hole_dim_err_mm=round(hole_dim_err, 4),
            holes_matched=holes_matched,
            holes_expected=len(plate.holes),
            passed=passed,
            issues=issues,
        ),
        outer_ring,
    )


def _merge_intervals_1d(intervals: list[tuple[float, float]]) -> list[tuple[float, float]]:
    if not intervals:
        return []
    ordered = sorted(intervals, key=lambda t: (t[0], t[1]))
    merged: list[list[float]] = [[ordered[0][0], ordered[0][1]]]
    for a, b in ordered[1:]:
        if a <= merged[-1][1] + 1e-7:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    return [(m[0], m[1]) for m in merged]


def _y_intervals_at_x(rings: list[Ring], x_mid: float) -> list[tuple[float, float]]:
    ivs: list[tuple[float, float]] = []
    for ring in rings:
        n = len(ring)
        if n < 3:
            continue
        ys_cross: list[float] = []
        for i in range(n):
            x1, y1 = ring[i]
            x2, y2 = ring[(i + 1) % n]
            lo_x, hi_x = (x1, x2) if x1 < x2 else (x2, x1)
            if lo_x < x_mid < hi_x:
                ys_cross.append(y1 + (x_mid - x1) * (y2 - y1) / (x2 - x1))
        ys_cross.sort()
        for j in range(0, len(ys_cross) - 1, 2):
            if ys_cross[j + 1] - ys_cross[j] > 1e-7:
                ivs.append((ys_cross[j], ys_cross[j + 1]))
    return _merge_intervals_1d(ivs)


def _rings_x_slab_metrics(
    rings_a: list[Ring],
    rings_b: list[Ring],
) -> tuple[float, float, float]:
    """Exact 2D (inter_area, union_area, iou_pct) via X-slab sweep with edge crossings."""
    segs: list[tuple[float, float, float, float]] = []
    xs_set: list[float] = []
    for r in list(rings_a) + list(rings_b):
        n = len(r)
        for i in range(n):
            x1, y1 = r[i]
            x2, y2 = r[(i + 1) % n]
            xs_set.append(x1)
            if abs(x2 - x1) > 1e-6:
                segs.append((x1, y1, x2, y2))

    if len(xs_set) < 2:
        return 0.0, 0.0, 0.0

    # Include intersection X points between non-vertical segments so crossing slopes are exact
    for i in range(len(segs)):
        ax1, ay1, ax2, ay2 = segs[i]
        ma_x = max(ax1, ax2)
        mi_x = min(ax1, ax2)
        for j in range(i + 1, len(segs)):
            bx1, by1, bx2, by2 = segs[j]
            if min(bx1, bx2) >= ma_x or max(bx1, bx2) <= mi_x:
                continue
            denom = (ax2 - ax1) * (by2 - by1) - (ay2 - ay1) * (bx2 - bx1)
            if abs(denom) < 1e-9:
                continue
            t = ((bx1 - ax1) * (by2 - by1) - (by1 - ay1) * (bx2 - bx1)) / denom
            u = ((bx1 - ax1) * (ay2 - ay1) - (by1 - ay1) * (ax2 - ax1)) / denom
            if 1e-6 < t < 1.0 - 1e-6 and 1e-6 < u < 1.0 - 1e-6:
                xs_set.append(ax1 + t * (ax2 - ax1))

    xs_sorted = sorted(xs_set)
    xs: list[float] = [xs_sorted[0]]
    for x in xs_sorted[1:]:
        if x - xs[-1] > 1e-5:
            xs.append(x)

    inter_area = 0.0
    union_area = 0.0
    for idx in range(len(xs) - 1):
        xl, xr = xs[idx], xs[idx + 1]
        dx = xr - xl
        if dx <= 1e-5:
            continue
        xm = 0.5 * (xl + xr)
        iv_a = _y_intervals_at_x(rings_a, xm)
        iv_b = _y_intervals_at_x(rings_b, xm)
        if not iv_a and not iv_b:
            continue
        len_a = sum(b - a for a, b in iv_a)
        len_b = sum(b - a for a, b in iv_b)
        len_i = 0.0
        ia = ib = 0
        while ia < len(iv_a) and ib < len(iv_b):
            lo = max(iv_a[ia][0], iv_b[ib][0])
            hi = min(iv_a[ia][1], iv_b[ib][1])
            if hi > lo:
                len_i += hi - lo
            if iv_a[ia][1] < iv_b[ib][1]:
                ia += 1
            else:
                ib += 1
        inter_area += dx * len_i
        union_area += dx * (len_a + len_b - len_i)

    iou_pct = (100.0 * inter_area / union_area) if union_area > 1e-6 else (100.0 if inter_area <= 1e-6 else 0.0)
    return inter_area, union_area, iou_pct


def _x_span_at_y(ring: Ring, y_target: float) -> Optional[tuple[float, float]]:
    """Return (min_x, max_x) where a polygon edge or horizontal scanline intersects y_target."""
    if not ring:
        return None
    xs: list[float] = []
    n = len(ring)
    for i in range(n):
        x1, y1 = ring[i]
        x2, y2 = ring[(i + 1) % n]
        if abs(y1 - y_target) <= 0.1 and abs(y2 - y_target) <= 0.1:
            xs.extend([x1, x2])
        elif (y1 <= y_target <= y2) or (y2 <= y_target <= y1):
            dy = y2 - y1
            if abs(dy) > 1e-9:
                xs.append(x1 + (y_target - y1) * (x2 - x1) / dy)
    return (min(xs), max(xs)) if xs else None


def _x_intervals_at_y(ring: Ring, y_target: float) -> list[tuple[float, float]]:
    """Return horizontal solid [x_lo, x_hi] intervals of a polygon clipped at scanline y_target (edge-inclusive)."""
    if len(ring) < 3:
        return []
    r0, r1, r2, r3 = polygon_bounds(ring)
    probe_y = y_target
    if abs(y_target - r1) <= 0.15:
        probe_y = r1 + min(0.25, 0.1 * max(1.0, r3 - r1))
    elif abs(y_target - r3) <= 0.15:
        probe_y = r3 - min(0.25, 0.1 * max(1.0, r3 - r1))
    xs: list[float] = []
    n = len(ring)
    for i in range(n):
        x1, y1 = ring[i]
        x2, y2 = ring[(i + 1) % n]
        lo_y, hi_y = (y1, y2) if y1 < y2 else (y2, y1)
        if lo_y < probe_y <= hi_y and abs(y2 - y1) > 1e-9:
            xs.append(x1 + (probe_y - y1) * (x2 - x1) / (y2 - y1))
    xs.sort()
    ivs: list[tuple[float, float]] = []
    for j in range(0, len(xs) - 1, 2):
        if xs[j + 1] - xs[j] > 0.1:
            ivs.append((xs[j], xs[j + 1]))
    return _merge_intervals_1d(ivs)


def _untransform_dxf_ring(
    dxf_ring: Ring,
    source_bounds: tuple[float, float, float, float],
    mirror_x: bool = False,
    mirror_y: bool = False,
) -> Ring:
    if not dxf_ring:
        return []
    d0, d1, d2, d3 = polygon_bounds(dxf_ring)
    dcx = 0.5 * (d0 + d2)
    dcy = 0.5 * (d1 + d3)
    unm = mirror_ring(dxf_ring, dcx, dcy, mirror_x=mirror_x, mirror_y=mirror_y)
    u0, u1, _, _ = polygon_bounds(unm)
    sx0, sy0, _, _ = source_bounds
    return clean_ring(translate_ring(unm, sx0 - u0, sy0 - u1), min_seg_len=0.05, ccw=True)


def simulate_3d_assembly(
    dstv: DSTVFile,
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
    plates: list[ExtractedPlate],
    re_read_rings: dict[str, Ring],
    dxf_audits: list[DrawnDxfAudit],
    locked_rings: dict[str, list[Ring]],
    l_gabarit: float,
    detected_type: str,
    talpa_side: str,
    bom_row: Optional[BomRow],
    applied_delta_l: float,
    mid_x_raw: float,
) -> Sim3DReport:
    """Virtually reassemble re-read DXF plates into 3D solids and audit vs locked .nc1 & BOM."""
    H = spec.total_height
    B = spec.total_width
    tw = spec.tw

    # 1. Recover 3D placements for every drawn plate
    # Determine web mounting Y-span: on Tip 1 / overhanging talpă, webs mount along the narrow capac bounds
    eff_talpa_side = talpa_side
    if detected_type in ("cheson_cu_talpa", "nominal_parametric") and eff_talpa_side not in ("bottom", "top"):
        u_rings = locked_rings.get("u", [])
        o_rings = locked_rings.get("o", [])
        dy_u = (max(polygon_bounds(r)[3] for r in u_rings) - min(polygon_bounds(r)[1] for r in u_rings)) if u_rings else spec.b_bot
        dy_o = (max(polygon_bounds(r)[3] for r in o_rings) - min(polygon_bounds(r)[1] for r in o_rings)) if o_rings else spec.b_top
        if dy_u > dy_o + 2.0 * tw + 1.0:
            eff_talpa_side = "bottom"
        elif dy_o > dy_u + 2.0 * tw + 1.0:
            eff_talpa_side = "top"

    y_box_lo, y_box_hi = 0.0, B
    if detected_type in ("cheson_cu_talpa", "nominal_parametric") and eff_talpa_side in ("bottom", "top"):
        narrow_face = "o" if eff_talpa_side == "bottom" else "u"
        narrow_rings = locked_rings.get(narrow_face, [])
        if narrow_rings:
            _, ny0, _, ny1 = polygon_bounds(narrow_rings[0])
            if ny1 - ny0 > 2.0 * tw + 1.0:
                y_box_lo, y_box_hi = ny0, ny1

    # Each solid entry: (plate, ring_2d, plane='XY'|'XZ', z_band=(z0, z1), y_band=(y0, y1), bbox3d=(x0,y0,z0,x1,y1,z1))
    solids: list[dict[str, Any]] = []

    for p in plates:
        dxf_r = re_read_rings.get(p.name, p.exterior)
        if p.face == "o":
            r_glob = _untransform_dxf_ring(dxf_r, p.source_bounds, mirror_x=False, mirror_y=cfg.mirror_flange_o_y)
            bx0, by0, bx1, by1 = polygon_bounds(r_glob)
            z0, z1 = max(0.0, H - p.thickness), H
            solids.append({
                "plate": p,
                "ring": r_glob,
                "plane": "XY",
                "z_band": (z0, z1),
                "y_band": (by0, by1),
                "bbox3d": (bx0, by0, z0, bx1, by1, z1),
            })
        elif p.face == "u":
            r_glob = _untransform_dxf_ring(dxf_r, p.source_bounds, mirror_x=False, mirror_y=cfg.mirror_flange_u_y)
            bx0, by0, bx1, by1 = polygon_bounds(r_glob)
            z0, z1 = 0.0, p.thickness
            solids.append({
                "plate": p,
                "ring": r_glob,
                "plane": "XY",
                "z_band": (z0, z1),
                "y_band": (by0, by1),
                "bbox3d": (bx0, by0, z0, bx1, by1, z1),
            })
        elif p.face in ("v", "h"):
            mx = cfg.mirror_web_h_x if p.face == "h" else False
            r_glob = _untransform_dxf_ring(dxf_r, p.source_bounds, mirror_x=mx, mirror_y=False)
            bx0, bz0, bx1, bz1 = polygon_bounds(r_glob)
            y0, y1 = (y_box_hi - tw, y_box_hi) if p.face == "v" else (y_box_lo, y_box_lo + tw)
            solids.append({
                "plate": p,
                "ring": r_glob,
                "plane": "XZ",
                "z_band": (bz0, bz1),
                "y_band": (y0, y1),
                "bbox3d": (bx0, y0, bz0, bx1, y1, bz1),
            })

    if not solids:
        return Sim3DReport(
            l_3d_mm=0.0, b_3d_mm=0.0, h_3d_mm=0.0,
            err_l_mm=l_gabarit, err_b_mm=B, err_h_mm=H,
            max_interference_mm=0.0, interference_pairs=[],
            max_seam_gap_mm=0.0, seam_notes=[],
            face_iou_pct={}, overall_3d_similarity_pct=0.0,
            holes_3d_expected=0, holes_3d_matched=0,
            max_hole_3d_err_mm=0.0, hole_shape_mismatches=0,
            mass_3d_kg=0.0, mass_bom_kg=(bom_row.weight_unit_kg if bom_row else None),
            mass_diff_pct=None, passed=False, issues=["Nicio placă pentru simulare 3D"],
        )

    # 2. Global 3D envelope
    min_x = min(s["bbox3d"][0] for s in solids)
    min_y = min(s["bbox3d"][1] for s in solids)
    min_z = min(s["bbox3d"][2] for s in solids)
    max_x = max(s["bbox3d"][3] for s in solids)
    max_y = max(s["bbox3d"][4] for s in solids)
    max_z = max(s["bbox3d"][5] for s in solids)

    l_3d = max_x - min_x
    b_3d = max_y - min_y
    h_3d = max_z - min_z
    err_l = abs(l_3d - l_gabarit)
    err_b = abs(b_3d - B)
    err_h = abs(h_3d - H)

    issues: list[str] = []
    if err_l > 0.01:
        issues.append(f"Lungime 3D {l_3d:.3f} != țintă {l_gabarit:.3f} mm")
    if err_b > 0.01:
        issues.append(f"Lățime 3D {b_3d:.3f} != profil {B:.3f} mm")
    if err_h > 0.01:
        issues.append(f"Înălțime 3D {h_3d:.3f} != profil {H:.3f} mm")

    # 3. Pairwise 3D collision / interference check (solid interiors must not penetrate > 0.01 mm)
    max_inter = 0.0
    inter_pairs: list[str] = []
    for i in range(len(solids)):
        ax0, ay0, az0, ax1, ay1, az1 = solids[i]["bbox3d"]
        for j in range(i + 1, len(solids)):
            bx0, by0, bz0, bx1, by1, bz1 = solids[j]["bbox3d"]
            ox = min(ax1, bx1) - max(ax0, bx0)
            oy = min(ay1, by1) - max(ay0, by0)
            oz = min(az1, bz1) - max(az0, bz0)
            if ox > 0.01 and oy > 0.01 and oz > 0.01:
                if solids[i]["plane"] == solids[j]["plane"]:
                    i_area, _, _ = _rings_x_slab_metrics([solids[i]["ring"]], [solids[j]["ring"]])
                    if i_area <= 0.1:
                        continue
                pen = min(ox, oy, oz)
                max_inter = max(max_inter, pen)
                inter_pairs.append(f"{solids[i]['plate'].name}↔{solids[j]['plate'].name} ({pen:.3f} mm)")

    if max_inter > 0.01:
        issues.append(f"Coliziune 3D detectată ({max_inter:.3f} mm): {', '.join(inter_pairs)}")

    # 4. Weld seam alignment checks (flanges <-> webs)
    flange_solids = [s for s in solids if s["plane"] == "XY"]
    web_solids = [s for s in solids if s["plane"] == "XZ"]
    max_seam_gap = 0.0
    seam_notes: list[str] = []

    for ws in web_solids:
        wp: ExtractedPlate = ws["plate"]
        wy_mid = 0.5 * (ws["y_band"][0] + ws["y_band"][1])
        wz0, wz1 = ws["z_band"]
        wz_mid = 0.5 * (wz0 + wz1)
        wx0, _, _, wx1, _, _ = ws["bbox3d"]
        for fs in flange_solids:
            fp: ExtractedPlate = fs["plate"]
            fz0, fz1 = fs["z_band"]
            fy0, fy1 = fs["y_band"]
            # Check transverse contact gap (flush at Y or Z)
            if fy0 - 0.5 <= ws["y_band"][1] <= fy1 + 0.5 and fy0 - 0.5 <= ws["y_band"][0] <= fy1 + 0.5:
                if detected_type == "cheson_dublu" and (
                    (fp.face == "o" and wz_mid < 0.5 * H)
                    or (fp.face == "u" and wz_mid > 0.5 * H)
                ):
                    continue
                # Over-web flange (e.g. talpa): Z should touch web top/bot
                gap_z = min(abs(wz0 - fz1), abs(fz0 - wz1))
                max_seam_gap = max(max_seam_gap, gap_z)
            elif (ws["y_band"][1] <= fy0 + 0.5) or (ws["y_band"][0] >= fy1 - 0.5):
                # Only compare vertically adjacent capac <-> web pairs (on Tip 3, bottom webs don't touch top capac)
                if min(wz1, fz1) <= max(wz0, fz0) - 0.5:
                    continue
                # Between-web capac: Y face touches web inner wall
                gap_y = abs(fy0 - ws["y_band"][1]) if ws["y_band"][1] <= fy0 + 0.5 else abs(ws["y_band"][0] - fy1)
                max_seam_gap = max(max_seam_gap, gap_y)
                # Also check X alignment along contact Y line if web wasn't shortened by cutout residuals
                if wp.removed_residuals == 0:
                    contact_y = fy0 if wp.face == "h" else fy1
                    f_span = _x_span_at_y(fs["ring"], contact_y)
                    w_z_ref = wz1 if fp.face == "o" else wz0
                    w_span = _x_span_at_y(ws["ring"], w_z_ref)
                    if f_span is not None and w_span is not None:
                        # Check if web spans at least flange contact (allowing outer chamfer/overhang)
                        if w_span[0] > f_span[0] + 0.05:
                            gap_x = w_span[0] - f_span[0]
                            max_seam_gap = max(max_seam_gap, gap_x)
                            seam_notes.append(f"{wp.name}/{fp.name} stânga +{gap_x:.3f}mm")
                        if w_span[1] < f_span[1] - 0.05:
                            gap_x = f_span[1] - w_span[1]
                            max_seam_gap = max(max_seam_gap, gap_x)
                            seam_notes.append(f"{wp.name}/{fp.name} dreapta +{gap_x:.3f}mm")

    if max_seam_gap > 0.05:
        issues.append(f"Rost/decalaj sudură talpă-laterală: {max_seam_gap:.3f} mm")

    # 5. 4-face 3D projection silhouette IoU vs locked .nc1 faces
    recessed_faces, _, _ = _detect_tip1_recessed_web_faces(
        locked_rings, spec, detected_type, eff_talpa_side
    )
    face_iou: dict[str, float] = {}
    for face_xy in ("o", "u"):
        nc_rings = locked_rings.get(face_xy, [])
        if not nc_rings:
            continue
        proj_xy: list[Ring] = [s["ring"] for s in flange_solids if s["plate"].face == face_xy]
        for ws in web_solids:
            wy0, wy1 = ws["y_band"]
            # Check web footprint at outer top/bot band
            span_x = _x_span_at_y(ws["ring"], ws["z_band"][1] if face_xy == "o" else ws["z_band"][0])
            if span_x is not None and span_x[1] - span_x[0] > 0.1:
                # Only project top-half / bottom-half web belonging to this outer face (for Tip 3 multi-band)
                wz_mid = 0.5 * (ws["z_band"][0] + ws["z_band"][1])
                if detected_type == "cheson_dublu":
                    if face_xy == "o" and wz_mid < 0.5 * H:
                        continue
                    if face_xy == "u" and wz_mid > 0.5 * H:
                        continue
                if detected_type in ("cheson_cu_talpa", "nominal_parametric") and (
                    (eff_talpa_side == "bottom" and face_xy == "u")
                    or (eff_talpa_side == "top" and face_xy == "o")
                ):
                    continue
                proj_xy.append([(span_x[0], wy0), (span_x[1], wy0), (span_x[1], wy1), (span_x[0], wy1)])
        _, _, iou = _rings_x_slab_metrics(proj_xy, nc_rings)
        face_iou[face_xy] = round(iou, 2)

    for face_xz in ("v", "h"):
        nc_rings = locked_rings.get(face_xz, [])
        if not nc_rings:
            continue
        if face_xz in recessed_faces:
            # On recessed Tip 1 overhanging-talpă faces, AK only contains the outer talpă edge strip
            proj_xz: list[Ring] = []
            y_probe = B if face_xz == "v" else 0.0
            for fs in flange_solids:
                if fs["plate"].role != "talpa":
                    continue
                fz0, fz1 = fs["z_band"]
                for x0, x1 in _x_intervals_at_y(fs["ring"], y_probe):
                    proj_xz.append([(x0, fz0), (x1, fz0), (x1, fz1), (x0, fz1)])
        else:
            proj_xz = [s["ring"] for s in web_solids if s["plate"].face == face_xz]
            y_probe = (B - tw) if face_xz == "v" else tw
            for fs in flange_solids:
                fz0, fz1 = fs["z_band"]
                fy0, fy1 = fs["y_band"]
                yp = min(max(y_probe, fy0), fy1)
                for x0, x1 in _x_intervals_at_y(fs["ring"], yp):
                    proj_xz.append([(x0, fz0), (x1, fz0), (x1, fz1), (x0, fz1)])
        _, _, iou = _rings_x_slab_metrics(proj_xz, nc_rings)
        face_iou[face_xz] = round(iou, 2)

    overall_iou = round(sum(face_iou.values()) / len(face_iou), 2) if face_iou else 100.0

    # Completeness guard for compound box beams (must never silently miss webs/flanges)
    min_webs = 4 if detected_type == "cheson_dublu" else 2
    min_flanges = 2
    missing_structural = False
    if len(web_solids) < min_webs:
        missing_structural = True
        issues.append(f"Lipsesc laterale (L): detectate {len(web_solids)} din minim {min_webs}!")
    if len(flange_solids) < min_flanges:
        missing_structural = True
        issues.append(f"Lipsesc tălpi/capace (T/C): detectate {len(flange_solids)} din minim {min_flanges}!")

    # 6. 3D Hole/Slot Cross-Check vs DSTV BO
    audit_by_name = {a.plate_name: a for a in dxf_audits}
    drawn_3d_holes: list[tuple[str, str, float, float, float, float, float]] = []
    for s in solids:
        p = s["plate"]
        aud = audit_by_name.get(p.name)
        if not aud:
            continue
        dxf_r = re_read_rings.get(p.name, p.exterior)
        d0, d1, d2, d3 = polygon_bounds(dxf_r)
        dcx, dcy = 0.5 * (d0 + d2), 0.5 * (d1 + d3)
        sx0, sy0, _, _ = p.source_bounds
        r_glob = s["ring"]
        gx0, gy0, _, _ = polygon_bounds(r_glob)
        mir_x = (cfg.mirror_web_h_x if p.face == "h" else False)
        mir_y = (cfg.mirror_flange_o_y if p.face == "o" else (cfg.mirror_flange_u_y if p.face == "u" else False))
        # Un-mirror & translate from dxf_r origin to global plate origin
        for dh in aud.holes_drawn:
            ux = (2.0 * dcx - dh.cx) if mir_x else dh.cx
            uy = (2.0 * dcy - dh.cy) if mir_y else dh.cy
            # Offset from unmirrored dxf bbox min
            unm_x = ux - (d0 if not mir_x else d0)
            unm_y = uy - (d1 if not mir_y else d1)
            gx = sx0 + (ux - d0)
            gy = sy0 + (uy - d1)
            drawn_3d_holes.append((p.face, dh.kind, gx, gy, dh.width_mm, dh.length_mm, dh.angle_deg))

    expected_3d_holes: list[tuple[str, str, float, float, float, float, float]] = []
    for f, hlist in dstv.holes.items():
        for hole in hlist:
            ekind, ecx, ecy, ew, el, eang = _expected_hole_geom(hole)
            if abs(applied_delta_l) > 1e-7 and ecx > mid_x_raw:
                ecx += applied_delta_l
            expected_3d_holes.append((f, ekind, ecx, ecy, ew, el, eang))

    holes_3d_matched = 0
    max_hole_3d_err = 0.0
    hole_shape_mismatches = 0
    used_3d: set[int] = set()
    for ef, ekind, ecx, ecy, ew, el, _ in expected_3d_holes:
        best_i = None
        best_d = 1e9
        for i_d, (df, dkind, dgx, dgy, dw, dl, _) in enumerate(drawn_3d_holes):
            if i_d in used_3d or df != ef:
                continue
            dist = math.hypot(dgx - ecx, dgy - ecy)
            if dist < best_d:
                best_d = dist
                best_i = i_d
        if best_i is None or best_d > 1.0:
            max_hole_3d_err = max(max_hole_3d_err, best_d if best_i is not None else 999.0)
            continue
        used_3d.add(best_i)
        _, dkind, _, _, dw, dl, _ = drawn_3d_holes[best_i]
        dim_e = max(abs(dw - ew), abs(dl - el))
        err_total = max(best_d, dim_e)
        max_hole_3d_err = max(max_hole_3d_err, err_total)
        if dkind != ekind:
            hole_shape_mismatches += 1
        elif err_total <= 0.05:
            holes_3d_matched += 1

    # 7. Net 3D steel mass vs STYKKLISTE BOM VEKT/STK (7.85 kg/dm3 = 7.85e-6 kg/mm3)
    mass_3d = round(sum(a.net_area_mm2 * a.thickness_mm * 7.85e-6 for a in dxf_audits), 2)
    mass_bom = bom_row.weight_unit_kg if (bom_row and bom_row.weight_unit_kg) else None
    mass_diff_pct = (
        round(abs(mass_3d - mass_bom) / mass_bom * 100.0, 2)
        if (mass_bom and mass_bom > 0.0)
        else None
    )
    mass_ok = True
    if mass_diff_pct is not None and mass_diff_pct > 10.0:
        mass_ok = False
        issues.append(
            f"Diferență mare masă 3D vs STYKKLISTE PDF: {mass_3d:.1f} kg vs {mass_bom:.1f} kg "
            f"({mass_diff_pct:.1f}% > 10%) — verificați piese lipsă!"
        )

    sim_passed = (
        not missing_structural
        and mass_ok
        and err_l <= 0.001 + 1e-6
        and err_b <= 0.001 + 1e-6
        and err_h <= 0.001 + 1e-6
        and max_inter <= 0.01
        and max_seam_gap <= 0.05
        and holes_3d_matched == len(expected_3d_holes)
        and hole_shape_mismatches == 0
        and max_hole_3d_err <= 0.05
        and overall_iou >= 95.0
    )

    return Sim3DReport(
        l_3d_mm=round(l_3d, 4),
        b_3d_mm=round(b_3d, 4),
        h_3d_mm=round(h_3d, 4),
        err_l_mm=round(err_l, 4),
        err_b_mm=round(err_b, 4),
        err_h_mm=round(err_h, 4),
        max_interference_mm=round(max_inter, 4),
        interference_pairs=inter_pairs,
        max_seam_gap_mm=round(max_seam_gap, 4),
        seam_notes=seam_notes,
        face_iou_pct=face_iou,
        overall_3d_similarity_pct=overall_iou,
        holes_3d_expected=len(expected_3d_holes),
        holes_3d_matched=holes_3d_matched,
        max_hole_3d_err_mm=round(max_hole_3d_err, 4),
        hole_shape_mismatches=hole_shape_mismatches,
        mass_3d_kg=mass_3d,
        mass_bom_kg=mass_bom,
        mass_diff_pct=mass_diff_pct,
        passed=sim_passed,
        issues=issues,
    )


_PDF_SCALE_TOKEN_RE = re.compile(r"^1\s*:\s*(\d+(?:\.\d+)?)$")


def _cubic_bezier_samples(p0, p1, p2, p3, steps: int = 6) -> list[tuple[float, float]]:
    pts: list[tuple[float, float]] = []
    for i in range(steps + 1):
        t = i / steps
        mt = 1.0 - t
        x = mt**3 * p0.x + 3.0 * mt**2 * t * p1.x + 3.0 * mt * t**2 * p2.x + t**3 * p3.x
        y = mt**3 * p0.y + 3.0 * mt**2 * t * p1.y + 3.0 * mt * t**2 * p2.y + t**3 * p3.y
        pts.append((float(x), float(y)))
    return pts


def _point_to_seg_dist_2d(px: float, py: float, x1: float, y1: float, x2: float, y2: float) -> float:
    dx, dy = x2 - x1, y2 - y1
    l2 = dx * dx + dy * dy
    if l2 <= 1e-12:
        return math.hypot(px - x1, py - y1)
    t = max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / l2))
    return math.hypot(px - (x1 + t * dx), py - (y1 + t * dy))


def _build_seg_spatial_grid(
    segs: list[tuple[float, float, float, float]],
    cell: float = 8.0,
) -> tuple[dict[tuple[int, int], list[int]], float]:
    grid: dict[tuple[int, int], list[int]] = {}
    for idx, (x1, y1, x2, y2) in enumerate(segs):
        ix0 = int(math.floor(min(x1, x2) / cell))
        ix1 = int(math.floor(max(x1, x2) / cell))
        iy0 = int(math.floor(min(y1, y2) / cell))
        iy1 = int(math.floor(max(y1, y2) / cell))
        for ix in range(ix0, ix1 + 1):
            for iy in range(iy0, iy1 + 1):
                grid.setdefault((ix, iy), []).append(idx)
    return grid, cell


def _extract_pdf_orthographic_views(pdf_path: Optional[str | Path]) -> Optional[dict[str, Any]]:
    """Extract TOP and FRONT true-scale dark vector segments & scales from a Tekla A-drawing PDF."""
    if pymupdf is None or not pdf_path:
        return None
    path = Path(pdf_path)
    if not path.is_file():
        return None

    doc = pymupdf.open(str(path))
    try:
        if len(doc) == 0:
            return None
        pg = doc[0]
        words = pg.get_text("words") or []
        views: dict[str, dict[str, Any]] = {}
        for vname in ("TOP", "FRONT"):
            for w in words:
                if str(w[4]).strip().upper() == vname:
                    cx = 0.5 * (float(w[0]) + float(w[2]))
                    cy = float(w[3])
                    sc = 25.0
                    for w2 in words:
                        m = _PDF_SCALE_TOKEN_RE.match(str(w2[4]).strip())
                        if m and abs(0.5 * (float(w2[0]) + float(w2[2])) - cx) < 80.0 and -5.0 < float(w2[1]) - cy < 45.0:
                            val = float(m.group(1))
                            if val >= 1.0:
                                sc = val
                                break
                    views[vname] = {
                        "bbox": (float(w[0]), float(w[1]), float(w[2]), float(w[3])),
                        "scale": sc,
                    }
                    break

        if "TOP" not in views and "FRONT" not in views:
            return None

        top_ymax = views["TOP"]["bbox"][1] if "TOP" in views else pg.rect.height * 0.35
        front_ymin = (views["TOP"]["bbox"][3] + 8.0) if "TOP" in views else pg.rect.height * 0.35
        front_ymax = views["FRONT"]["bbox"][1] if "FRONT" in views else pg.rect.height * 0.65

        top_segs: list[tuple[float, float, float, float]] = []
        front_segs: list[tuple[float, float, float, float]] = []
        for d in pg.get_drawings():
            col = d.get("color")
            if col is None or max(col) > 0.25:
                continue
            for it in d.get("items", []):
                kind = it[0]
                sub: list[tuple[float, float, float, float]] = []
                if kind == "l":
                    p1, p2 = it[1], it[2]
                    if math.hypot(p2.x - p1.x, p2.y - p1.y) >= 0.15:
                        sub.append((float(p1.x), float(p1.y), float(p2.x), float(p2.y)))
                elif kind == "c":
                    pts = _cubic_bezier_samples(it[1], it[2], it[3], it[4], steps=6)
                    for k in range(len(pts) - 1):
                        a, b = pts[k], pts[k + 1]
                        if math.hypot(b[0] - a[0], b[1] - a[1]) >= 0.12:
                            sub.append((a[0], a[1], b[0], b[1]))
                elif kind == "re":
                    r = it[1]
                    sub.extend([
                        (float(r.x0), float(r.y0), float(r.x1), float(r.y0)),
                        (float(r.x1), float(r.y0), float(r.x1), float(r.y1)),
                        (float(r.x1), float(r.y1), float(r.x0), float(r.y1)),
                        (float(r.x0), float(r.y1), float(r.x0), float(r.y0)),
                    ])
                for seg in sub:
                    ym = 0.5 * (seg[1] + seg[3])
                    if 25.0 < ym < top_ymax:
                        top_segs.append(seg)
                    elif front_ymin < ym < front_ymax:
                        front_segs.append(seg)

        top_grid, cell = _build_seg_spatial_grid(top_segs, cell=8.0) if top_segs else ({}, 8.0)
        front_grid, _ = _build_seg_spatial_grid(front_segs, cell=8.0) if front_segs else ({}, 8.0)
        return {
            "pdf_name": path.name,
            "top_scale": views["TOP"]["scale"] if "TOP" in views else None,
            "front_scale": views["FRONT"]["scale"] if "FRONT" in views else None,
            "top_segs": top_segs,
            "front_segs": front_segs,
            "top_grid": top_grid,
            "front_grid": front_grid,
            "cell": cell,
        }
    finally:
        doc.close()


def _build_assembled_pdf_projections(
    plates: list[ExtractedPlate],
    re_read_rings: dict[str, Ring],
    dxf_audits: list[DrawnDxfAudit],
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
    detected_type: str,
) -> tuple[list[tuple[float, float, float, float]], list[tuple[float, float, float, float]], list[tuple[str, float, float, float]]]:
    """Build 3D-assembled TOP (XY), FRONT (XZ), and hole center projections from re-read DXFs."""
    H, B, tw = spec.total_height, spec.total_width, spec.tw
    y_box_lo, y_box_hi = 0.0, B
    if detected_type == "cheson_cu_talpa":
        cap = next((p for p in plates if p.role == "capac"), None)
        if cap:
            y_box_lo, y_box_hi = cap.source_bounds[1] - tw, cap.source_bounds[3] + tw

    top_segs: list[tuple[float, float, float, float]] = []
    front_segs: list[tuple[float, float, float, float]] = []
    audit_by_name = {a.plate_name: a for a in dxf_audits}
    holes_proj: list[tuple[str, float, float, float]] = []  # ("TOP"|"FRONT", gx, gy_or_gz, r_mm)

    for p in plates:
        dxf_r = re_read_rings.get(p.name, p.exterior)
        aud = audit_by_name.get(p.name)
        mir_x = cfg.mirror_web_h_x if p.face == "h" else False
        mir_y = cfg.mirror_flange_o_y if p.face == "o" else (cfg.mirror_flange_u_y if p.face == "u" else False)
        if p.face in ("o", "u"):
            rg = _untransform_dxf_ring(dxf_r, p.source_bounds, mirror_x=False, mirror_y=mir_y)
            is_between_capac = p.role in ("capac", "flange_top", "flange_bot") and abs((p.source_bounds[3] - p.source_bounds[1]) - ((y_box_hi - y_box_lo) - 2.0 * tw)) <= 1.0
            for i in range(len(rg)):
                a, b = rg[i], rg[(i + 1) % len(rg)]
                ay, by = a[1], b[1]
                # Map inner-wall horizontal edges of between-webs capac onto outer assembled beam edges;
                # actual inner tw seam lines come strictly from the web segments where webs exist.
                if is_between_capac and abs(ay - by) <= 0.2:
                    if abs(ay - (y_box_lo + tw)) <= 0.5:
                        ay = by = y_box_lo
                    elif abs(ay - (y_box_hi - tw)) <= 0.5:
                        ay = by = y_box_hi
                top_segs.append((a[0], ay, b[0], by))
            z0, z1 = (H - p.thickness, H) if p.face == "o" else (0.0, p.thickness)
            bx0, _, bx1, _ = polygon_bounds(rg)
            front_segs.extend([(bx0, z0, bx1, z0), (bx0, z1, bx1, z1)])
        elif p.face in ("v", "h"):
            rg = _untransform_dxf_ring(dxf_r, p.source_bounds, mirror_x=mir_x, mirror_y=False)
            rg_flush: Ring = []
            for x, z in rg:
                nz = z
                if abs(z - tw) <= 0.5 and detected_type != "cheson_cu_talpa":
                    nz = 0.0
                elif abs(z - (H - tw)) <= 0.5:
                    nz = H
                rg_flush.append((x, nz))
            for i in range(len(rg_flush)):
                a, b = rg_flush[i], rg_flush[(i + 1) % len(rg_flush)]
                front_segs.append((a[0], a[1], b[0], b[1]))
            y0, y1 = (y_box_hi - tw, y_box_hi) if p.face == "v" else (y_box_lo, y_box_lo + tw)
            _, bz0, _, bz1 = polygon_bounds(rg)
            seen_iv: set[tuple[float, float]] = set()
            for z_probe in (bz1, bz0):
                for sx0, sx1 in _x_intervals_at_y(rg, z_probe):
                    k_iv = (round(sx0, 1), round(sx1, 1))
                    if k_iv not in seen_iv:
                        seen_iv.add(k_iv)
                        top_segs.extend([(sx0, y0, sx1, y0), (sx0, y1, sx1, y1)])

        if aud and aud.holes_drawn and p.face in ("o", "u", "v", "h"):
            d0, d1, d2, d3 = polygon_bounds(dxf_r)
            dcx, dcy = 0.5 * (d0 + d2), 0.5 * (d1 + d3)
            sx0, sy0, _, _ = p.source_bounds
            vname = "TOP" if p.face in ("o", "u") else "FRONT"
            for dh in aud.holes_drawn:
                ux = (2.0 * dcx - dh.cx) if mir_x else dh.cx
                uy = (2.0 * dcy - dh.cy) if mir_y else dh.cy
                gx = sx0 + (ux - d0)
                gy = sy0 + (uy - d1)
                holes_proj.append((vname, gx, gy, 0.5 * max(dh.width_mm, dh.length_mm)))

    return top_segs, front_segs, holes_proj


def _eval_view_overlay(
    model_segs_mm: list[tuple[float, float, float, float]],
    pdf_segs: list[tuple[float, float, float, float]],
    grid: dict[tuple[int, int], list[int]],
    cell: float,
    scale: float,
    ox_pt: float,
    oy_pt: float,
    flip_y: bool,
    span_y_mm: float,
    flip_x: bool,
    span_x_mm: float,
    tol_pt: float = 0.25,
) -> tuple[float, float, float, float]:
    """Return (overlay_pct, mean_err_pt, matched_len_mm, total_len_mm)."""
    s = (72.0 / 25.4) / scale
    total_len = 0.0
    matched_len = 0.0
    sum_err = 0.0

    for ax, ay, bx, by in model_segs_mm:
        seg_len = math.hypot(bx - ax, by - ay)
        if seg_len < 0.5:
            continue
        n_samples = max(2, min(40, int(seg_len / 50.0) + 1))
        step_w = seg_len / n_samples
        for k in range(n_samples):
            t = (k + 0.5) / n_samples
            mx = ax + t * (bx - ax)
            my = ay + t * (by - ay)
            if flip_x:
                mx = span_x_mm - mx
            if flip_y:
                my = span_y_mm - my
            px = ox_pt + mx * s
            py = oy_pt + my * s
            ix, iy = int(math.floor(px / cell)), int(math.floor(py / cell))
            best = 1e9
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for sidx in grid.get((ix + dx, iy + dy), ()):
                        x1, y1, x2, y2 = pdf_segs[sidx]
                        d = _point_to_seg_dist_2d(px, py, x1, y1, x2, y2)
                        if d < best:
                            best = d
                            if best <= 0.02:
                                break
            total_len += step_w
            if best <= tol_pt:
                matched_len += step_w
                sum_err += best * step_w

    pct = (100.0 * matched_len / total_len) if total_len > 0.0 else 0.0
    mean_err = (sum_err / matched_len) if matched_len > 0.0 else 999.0
    return pct, mean_err, matched_len, total_len


def _register_view_to_pdf(
    model_segs_mm: list[tuple[float, float, float, float]],
    L_mm: float,
    W_mm: float,
    pdf_segs: list[tuple[float, float, float, float]],
    grid: dict[tuple[int, int], list[int]],
    cell: float,
    scale: float,
    allowed_fy: tuple[bool, ...] = (False, True),
    allowed_fx: tuple[bool, ...] = (False, True),
    hint_ox_pt: Optional[float] = None,
) -> tuple[float, float, float, float, Optional[tuple[float, float, bool, bool]]]:
    if not model_segs_mm or not pdf_segs or L_mm <= 1.0 or W_mm <= 1.0:
        return 0.0, 999.0, 0.0, 0.0, None

    s = (72.0 / 25.4) / scale
    L_pt, W_pt = L_mm * s, W_mm * s

    by_y: dict[float, list[tuple[float, float]]] = {}
    for x1, y1, x2, y2 in pdf_segs:
        if abs(y1 - y2) < 0.35 and abs(x2 - x1) > 4.0:
            ky = round(y1, 1)
            by_y.setdefault(ky, []).append((min(x1, x2), max(x1, x2)))

    y_runs: dict[float, list[tuple[float, float]]] = {}
    for ky, ivs in by_y.items():
        merged = _merge_intervals_1d([(a - 2.5, b + 2.5) for a, b in ivs])
        runs = [(a + 2.5, b - 2.5) for a, b in merged if (b - a) >= max(10.0, 0.25 * L_pt)]
        if runs:
            y_runs[ky] = runs

    ys_long = sorted(y_runs.keys())
    model_h_xs: set[float] = set()
    for ax, ay, bx, by in model_segs_mm:
        if abs(ay - by) < 1.0 and abs(bx - ax) > 0.30 * L_mm:
            model_h_xs.add(round(min(ax, bx), 1))
            model_h_xs.add(round(max(ax, bx), 1))
    model_x_offsets_pt = sorted({round(x * s, 2) for x in model_h_xs})[:10]

    cand_anchors: list[tuple[float, float]] = []
    for i, y1 in enumerate(ys_long):
        for y2 in ys_long[i + 1:]:
            if abs((y2 - y1) - W_pt) <= 1.5:
                xs_cand: set[float] = set()
                for ky in (y1, y2):
                    for a, b in y_runs[ky]:
                        xs_cand.add(round(a, 1))
                        xs_cand.add(round(b - L_pt, 1))
                        for off in model_x_offsets_pt:
                            xs_cand.add(round(a - off, 1))
                            xs_cand.add(round(b - off, 1))
                for ox in xs_cand:
                    if -50.0 < ox < 2200.0:
                        cand_anchors.append((ox, y1))
                        if abs((y2 - y1) - W_pt) > 0.2:
                            cand_anchors.append((ox, y2 - W_pt))

    if hint_ox_pt is not None:
        hint_ys: list[float] = []
        for _, oy in cand_anchors:
            if not any(abs(oy - hy) <= 0.25 for hy in hint_ys):
                hint_ys.append(oy)
        for hy in hint_ys:
            cand_anchors.append((hint_ox_pt, hy))

    uniq: list[tuple[float, float]] = []
    seen: set[tuple[int, int]] = set()
    for ox, oy in cand_anchors:
        k = (round(ox / 0.25), round(oy / 0.25))
        if k not in seen:
            seen.add(k)
            uniq.append((ox, oy))

    if not uniq:
        return 0.0, 999.0, 0.0, 0.0, None

    # Deduplicate model segments to 0.1 mm precision and preserve longest horizontal + non-horizontal bevel segments
    dedup_segs: list[tuple[float, float, float, float]] = []
    seen_segs: set[tuple[tuple[int, int], tuple[int, int]]] = set()
    for x1, y1, x2, y2 in model_segs_mm:
        p1 = (round(x1 * 10.0), round(y1 * 10.0))
        p2 = (round(x2 * 10.0), round(y2 * 10.0))
        sk = (p1, p2) if p1 <= p2 else (p2, p1)
        if sk not in seen_segs:
            seen_segs.add(sk)
            dedup_segs.append((x1, y1, x2, y2))

    h_segs = sorted(
        (g for g in dedup_segs if abs(g[3] - g[1]) <= 0.25),
        key=lambda g: -math.hypot(g[2] - g[0], g[3] - g[1]),
    )
    non_h_segs = sorted(
        (g for g in dedup_segs if abs(g[3] - g[1]) > 0.25),
        key=lambda g: -math.hypot(g[2] - g[0], g[3] - g[1]),
    )
    fast_segs = h_segs[:10] + non_h_segs[:16]
    ordered_segs = sorted(dedup_segs, key=lambda g: -math.hypot(g[2] - g[0], g[3] - g[1]))
    coarse_segs = list(dict.fromkeys(ordered_segs[:45] + non_h_segs[:16]))

    fast_scored: list[tuple[float, float, tuple[float, float, bool, bool]]] = []
    for ox, oy in uniq:
        for fy in allowed_fy:
            for fx in allowed_fx:
                pct, me, _, _ = _eval_view_overlay(
                    fast_segs, pdf_segs, grid, cell, scale, ox, oy, fy, W_mm, fx, L_mm, tol_pt=0.35
                )
                if pct >= 60.0:
                    fast_scored.append((pct, -me, (ox, oy, fy, fx)))
    fast_scored.sort(key=lambda item: (round(item[0], 1), item[1]), reverse=True)
    top_cands = [item[2] for item in fast_scored[:24]] if fast_scored else [(uniq[0][0], uniq[0][1], allowed_fy[0], allowed_fx[0])]
    if hint_ox_pt is not None:
        seen_cands = set(top_cands)
        for ox, oy in uniq:
            if abs(ox - hint_ox_pt) <= 0.3:
                for fy in allowed_fy:
                    for fx in allowed_fx:
                        cand = (ox, oy, fy, fx)
                        if cand not in seen_cands:
                            seen_cands.add(cand)
                            top_cands.append(cand)

    best_c: tuple[float, float, Optional[tuple[float, float, bool, bool]]] = (-1.0, 999.0, None)
    for ox, oy, fy, fx in top_cands:
        pct, me, _, _ = _eval_view_overlay(
            coarse_segs, pdf_segs, grid, cell, scale, ox, oy, fy, W_mm, fx, L_mm, tol_pt=0.28
        )
        if (pct > best_c[0] + 0.05) or (abs(pct - best_c[0]) <= 0.05 and me < best_c[1]):
            best_c = (pct, me, (ox, oy, fy, fx))

    if best_c[2] is None:
        return 0.0, 999.0, 0.0, 0.0, None

    ox0, oy0, fy, fx = best_c[2]
    # Stage 1 fine search (±0.8 pt at 0.12 pt step) -> catches any tw-shift or edge-stroke offset
    best_s1 = (-1.0, 999.0, ox0, oy0)
    steps_s1 = [-0.72, -0.48, -0.24, -0.12, 0.0, 0.12, 0.24, 0.48, 0.72]
    for dx in steps_s1:
        for dy in steps_s1:
            pct, me, _, _ = _eval_view_overlay(
                coarse_segs, pdf_segs, grid, cell, scale, ox0 + dx, oy0 + dy, fy, W_mm, fx, L_mm, tol_pt=0.25
            )
            if (pct > best_s1[0] + 1e-5) or (abs(pct - best_s1[0]) <= 1e-5 and me < best_s1[1]):
                best_s1 = (pct, me, ox0 + dx, oy0 + dy)

    # Stage 2 sub-point refinement (±0.12 pt at 0.04 pt step) over all model segments
    ox1, oy1 = best_s1[2], best_s1[3]
    best_f = (-1.0, 999.0, 0.0, 0.0, (ox1, oy1, fy, fx))
    for dx in (-0.12, -0.08, -0.04, 0.0, 0.04, 0.08, 0.12):
        for dy in (-0.12, -0.08, -0.04, 0.0, 0.04, 0.08, 0.12):
            pct, me, ml, tl = _eval_view_overlay(
                model_segs_mm, pdf_segs, grid, cell, scale, ox1 + dx, oy1 + dy, fy, W_mm, fx, L_mm, tol_pt=0.25
            )
            if (pct > best_f[0] + 1e-6) or (abs(pct - best_f[0]) <= 1e-6 and me < best_f[1]):
                best_f = (pct, me, ml, tl, (ox1 + dx, oy1 + dy, fy, fx))
    return best_f


def _count_uncovered_interior_web_lines(
    model_segs_mm: list[tuple[float, float, float, float]],
    L_mm: float,
    B_mm: float,
    pdf_segs: list[tuple[float, float, float, float]],
    scale: float,
    reg: tuple[float, float, bool, bool],
) -> int:
    """Detect long interior PDF lines (missed webs) inside the registered TOP beam footprint."""
    ox_pt, oy_pt, fy, fx = reg
    s = (72.0 / 25.4) / scale
    L_pt, B_pt = L_mm * s, B_mm * s
    by_y: dict[float, list[tuple[float, float]]] = {}
    for x1, y1, x2, y2 in pdf_segs:
        if abs(y1 - y2) < 0.25 and ox_pt - 1.0 <= min(x1, x2) and max(x1, x2) <= ox_pt + L_pt + 1.0:
            if oy_pt + 0.35 <= y1 <= oy_pt + B_pt - 0.35:
                by_y.setdefault(round(y1, 1), []).append((min(x1, x2), max(x1, x2)))

    # Project horizontal model lines into PDF coords
    m_hlines: list[tuple[float, float, float]] = []
    for ax, ay, bx, by in model_segs_mm:
        if abs(ay - by) < 0.5 and abs(bx - ax) > 0.20 * L_mm:
            mx0 = ox_pt + ((L_mm - max(ax, bx)) if fx else min(ax, bx)) * s
            mx1 = ox_pt + ((L_mm - min(ax, bx)) if fx else max(ax, bx)) * s
            my = oy_pt + ((B_mm - ay) if fy else ay) * s
            m_hlines.append((my, mx0, mx1))

    missing = 0
    for ky, ivs in by_y.items():
        merged = _merge_intervals_1d([(a - 0.8, b + 0.8) for a, b in ivs])
        for a, b in merged:
            if (b - a) < 0.70 * L_pt:
                continue
            mid_x = 0.5 * (a + b)
            if not any(abs(ky - my) <= 0.45 and mx0 - 2.0 <= mid_x <= mx1 + 2.0 for my, mx0, mx1 in m_hlines):
                missing += 1
    return missing


def verify_assembly_pdf_overlay(
    pdf_path: Optional[str | Path],
    plates: list[ExtractedPlate],
    re_read_rings: dict[str, Ring],
    dxf_audits: list[DrawnDxfAudit],
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
    detected_type: str,
    l_assembled_mm: float,
    cached_pdf_views: Optional[dict[str, Any]] = None,
) -> Optional[PdfOverlayReport]:
    """Scale-matched 2D orthographic overlay verification (>= 99.9%) vs A-drawing PDF TOP & FRONT views."""
    views_data = cached_pdf_views if cached_pdf_views is not None else _extract_pdf_orthographic_views(pdf_path)
    if views_data is None:
        return None

    top_segs_m, front_segs_m, holes_proj = _build_assembled_pdf_projections(
        plates, re_read_rings, dxf_audits, spec, cfg, detected_type
    )
    H, B, L = spec.total_height, spec.total_width, l_assembled_mm

    top_scale = views_data["top_scale"] or 25.0
    front_scale = views_data["front_scale"] or 25.0

    t_pct, t_me_pt, t_ml, t_tl, t_reg = _register_view_to_pdf(
        top_segs_m, L, B, views_data["top_segs"], views_data["top_grid"], views_data["cell"], top_scale
    )

    fx_allow: tuple[bool, ...] = (t_reg[3],) if (t_reg is not None and t_pct >= 98.0) else (False, True)
    front_hint_ox = t_reg[0] if (t_reg is not None and t_pct >= 98.0 and abs(top_scale - front_scale) <= 1e-6) else None
    f_pct, f_me_pt, f_ml, f_tl, f_reg = _register_view_to_pdf(
        front_segs_m, L, H, views_data["front_segs"], views_data["front_grid"], views_data["cell"], front_scale,
        allowed_fy=(True,), allowed_fx=fx_allow, hint_ox_pt=front_hint_ox,
    )
    if f_pct < 98.0 and len(fx_allow) == 1:
        f_pct, f_me_pt, f_ml, f_tl, f_reg = _register_view_to_pdf(
            front_segs_m, L, H, views_data["front_segs"], views_data["front_grid"], views_data["cell"], front_scale,
            allowed_fy=(True,), allowed_fx=(False, True), hint_ox_pt=front_hint_ox,
        )
    if t_pct < 99.9 and f_pct >= 99.9 and f_reg is not None and abs(top_scale - front_scale) <= 1e-6:
        t2_pct, t2_me_pt, t2_ml, t2_tl, t2_reg = _register_view_to_pdf(
            top_segs_m, L, B, views_data["top_segs"], views_data["top_grid"], views_data["cell"], top_scale,
            allowed_fx=(f_reg[3],), hint_ox_pt=f_reg[0],
        )
        if (t2_pct > t_pct + 1e-6) or (abs(t2_pct - t_pct) <= 1e-6 and t2_me_pt < t_me_pt):
            t_pct, t_me_pt, t_ml, t_tl, t_reg = t2_pct, t2_me_pt, t2_ml, t2_tl, t2_reg

    missing_lines = 0
    if t_reg is not None:
        missing_lines = _count_uncovered_interior_web_lines(
            top_segs_m, L, B, views_data["top_segs"], top_scale, t_reg
        )

    # Check projected holes on PDF TOP/FRONT views
    holes_exp = 0
    holes_hit = 0
    cell = views_data["cell"]
    for vname, hx, hy, r_mm in holes_proj:
        reg = t_reg if vname == "TOP" else f_reg
        sc = top_scale if vname == "TOP" else front_scale
        span_y = B if vname == "TOP" else H
        p_segs = views_data["top_segs"] if vname == "TOP" else views_data["front_segs"]
        p_grid = views_data["top_grid"] if vname == "TOP" else views_data["front_grid"]
        if reg is None or not p_segs:
            continue
        holes_exp += 1
        ox, oy, fy, fx = reg
        s = (72.0 / 25.4) / sc
        mx = (L - hx) if fx else hx
        my = (span_y - hy) if fy else hy
        px = ox + mx * s
        py = oy + my * s
        tol_h = max(0.35, r_mm * s + 0.45)
        ix, iy = int(math.floor(px / cell)), int(math.floor(py / cell))
        found_h = False
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for sidx in p_grid.get((ix + dx, iy + dy), ()):
                    x1, y1, x2, y2 = p_segs[sidx]
                    if _point_to_seg_dist_2d(px, py, x1, y1, x2, y2) <= tol_h:
                        found_h = True
                        break
                if found_h:
                    break
            if found_h:
                break
        if found_h:
            holes_hit += 1

    raw_tot = t_tl + f_tl
    raw_pct = (100.0 * (t_ml + f_ml) / raw_tot) if raw_tot > 0.0 else 0.0
    penalty = 15.0 * missing_lines
    overall_pct = round(max(0.0, raw_pct - penalty), 2)
    t_pct_pen = round(max(0.0, t_pct - penalty), 2)

    pt_to_mm = 25.4 / 72.0
    t_err_mm = round(t_me_pt * pt_to_mm * top_scale, 3) if t_me_pt < 900.0 else 999.0
    f_err_mm = round(f_me_pt * pt_to_mm * front_scale, 3) if f_me_pt < 900.0 else 999.0

    issues: list[str] = []
    if missing_lines > 0:
        issues.append(f"PDF TOP: {missing_lines} linii structurale interne neacoperite (posibil piese lipsă!)")
    if overall_pct < 99.9:
        issues.append(
            f"Suprapunere asamblată PDF sub prag ({overall_pct:.2f}% < 99.90%): "
            f"TOP={t_pct_pen:.2f}%, FRONT={f_pct:.2f}%"
        )

    passed = (overall_pct >= 99.9) and (missing_lines == 0)
    return PdfOverlayReport(
        pdf_name=views_data["pdf_name"],
        top_scale=views_data["top_scale"],
        front_scale=views_data["front_scale"],
        top_overlay_pct=t_pct_pen,
        front_overlay_pct=round(f_pct, 2),
        overall_overlay_pct=overall_pct,
        top_mean_err_mm=t_err_mm,
        front_mean_err_mm=f_err_mm,
        holes_pdf_expected=holes_exp,
        holes_pdf_matched=holes_hit,
        missing_structural_lines=missing_lines,
        passed=passed,
        issues=issues,
    )


def verify_post_draw_and_3d(
    dstv: DSTVFile,
    cfg: ChesonConfig,
    plates: list[ExtractedPlate],
    dxf_audits: list[DrawnDxfAudit],
    re_read_rings: dict[str, Ring],
    bom: Optional[BomTable] = None,
    bom_row: Optional[BomRow] = None,
    pdf_dims: Optional[set[float]] = None,
    pdf_analysis: Optional[PdfDrawingAnalysis] = None,
    a_pdf_path: Optional[Path] = None,
    cached_pdf_views: Optional[dict[str, Any]] = None,
) -> PostDrawVerificationReport:
    """Post-draw verification: re-read DXFs vs NC1 geometry, PDF overlay >=99.9%, and 3D simulation."""
    spec = parse_compound_profile(dstv.header.profile_str, dstv.header)
    spec = _align_spec_flanges_to_faces(dstv, spec, cfg)
    detected_type, talpa_side = detect_cheson_type(dstv, spec, cfg)

    raw_rings: dict[str, list[Ring]] = {}
    for face in ("o", "u", "v", "h"):
        raw_rings[face] = []
        for c in dstv.ak_contours.get(face, []):
            r = contour_to_ring(c, cfg)
            if len(r) >= 3 and abs(polygon_area(r)) > 0.1:
                raw_rings[face].append(r)

    x_min_env, _, l_nc1_raw = _compute_raw_x_envelope(raw_rings)
    l_gabarit, _ = _resolve_gabarit_length(l_nc1_raw, cfg)
    if cfg.exact_gabarit_lock:
        max_delta = 100.0 if (cfg.target_length_mm is not None and cfg.target_length_mm > 0.0) else 20.0
        locked_rings = _apply_gabarit_x_lock_to_rings(
            raw_rings, x_min_env, l_nc1_raw, l_gabarit, max_delta_mm=max_delta
        )
    else:
        locked_rings = raw_rings
    _, _, l_assembled_env = _compute_raw_x_envelope(locked_rings)
    mid_x_raw = x_min_env + 0.5 * l_nc1_raw
    applied_delta_l = (l_assembled_env - l_nc1_raw) if cfg.exact_gabarit_lock else 0.0
    l_sim_target = l_gabarit if cfg.exact_gabarit_lock else l_assembled_env

    sim_3d = simulate_3d_assembly(
        dstv=dstv,
        spec=spec,
        cfg=cfg,
        plates=plates,
        re_read_rings=re_read_rings,
        dxf_audits=dxf_audits,
        locked_rings=locked_rings,
        l_gabarit=l_sim_target,
        detected_type=detected_type,
        talpa_side=talpa_side,
        bom_row=bom_row,
        applied_delta_l=applied_delta_l,
        mid_x_raw=mid_x_raw,
    )

    all_dxf_clean = all(a.dxf_errors == 0 and a.outer_closed and a.passed for a in dxf_audits)
    all_layer_zero = all(a.layer_zero_ok for a in dxf_audits)
    nc1_max_dim_err = max((a.dim_err_mm for a in dxf_audits), default=0.0)
    nc1_max_hole_err = max(
        (max(a.hole_pos_err_mm, a.hole_dim_err_mm) for a in dxf_audits), default=0.0
    )

    pdf_bom_len_err: Optional[float] = None
    if bom_row and bom_row.length_mm:
        pdf_bom_len_err = round(abs(sim_3d.l_3d_mm - bom_row.length_mm), 4)

    # Confirm key dimensions appearing both in the PDF drawing and in our generated DXF geometry
    candidate_dims: set[float] = {
        round(sim_3d.l_3d_mm, 1),
        round(sim_3d.b_3d_mm, 1),
        round(sim_3d.h_3d_mm, 1),
        round(spec.h_web, 1),
        round(spec.b_bot, 1),
        round(spec.b_top, 1),
    }
    for a in dxf_audits:
        candidate_dims.add(round(a.length_mm, 1))
        candidate_dims.add(round(a.width_mm, 1))
    # Edge distances of holes / cutouts from plate borders and hole-to-hole pitches
    for a in dxf_audits:
        for h in a.holes_drawn:
            candidate_dims.add(round(h.cx, 1))
            candidate_dims.add(round(a.length_mm - h.cx, 1))
            candidate_dims.add(round(h.cy, 1))
            candidate_dims.add(round(a.width_mm - h.cy, 1))
            candidate_dims.add(round(h.width_mm, 1))
            candidate_dims.add(round(h.length_mm, 1))
        for i_h in range(len(a.holes_drawn)):
            for j_h in range(i_h + 1, len(a.holes_drawn)):
                dx_p = abs(a.holes_drawn[i_h].cx - a.holes_drawn[j_h].cx)
                dy_p = abs(a.holes_drawn[i_h].cy - a.holes_drawn[j_h].cy)
                if dx_p >= 10.0:
                    candidate_dims.add(round(dx_p, 1))
                if dy_p >= 10.0:
                    candidate_dims.add(round(dy_p, 1))
        rr = re_read_rings.get(a.plate_name, [])
        for px, py in rr:
            if 10.0 <= px <= a.length_mm - 10.0:
                candidate_dims.add(round(px, 1))
                candidate_dims.add(round(a.length_mm - px, 1))
        if sim_3d.l_3d_mm - a.length_mm >= 10.0:
            candidate_dims.add(round(sim_3d.l_3d_mm - a.length_mm, 1))
    for i_a in range(len(dxf_audits)):
        for j_a in range(i_a + 1, len(dxf_audits)):
            dl_p = abs(dxf_audits[i_a].length_mm - dxf_audits[j_a].length_mm)
            if dl_p >= 10.0:
                candidate_dims.add(round(dl_p, 1))

    eff_pdf_dims = pdf_dims if pdf_dims is not None else (pdf_analysis.dimensions if pdf_analysis else None)
    pdf_confirmed: list[float] = []
    if eff_pdf_dims:
        for cval in sorted(candidate_dims):
            if cval >= 10.0 and any(abs(cval - pd_val) <= 0.5 for pd_val in eff_pdf_dims):
                pdf_confirmed.append(round(cval, 2))

    pdf_overlay: Optional[PdfOverlayReport] = None
    eff_pdf_path = a_pdf_path or (bom.source_pdf if bom is not None else None)
    if eff_pdf_path is not None or cached_pdf_views is not None:
        pdf_overlay = verify_assembly_pdf_overlay(
            pdf_path=eff_pdf_path,
            plates=plates,
            re_read_rings=re_read_rings,
            dxf_audits=dxf_audits,
            spec=spec,
            cfg=cfg,
            detected_type=detected_type,
            l_assembled_mm=sim_3d.l_3d_mm if sim_3d.l_3d_mm > 0.0 else l_gabarit,
            cached_pdf_views=cached_pdf_views,
        )

    issues: list[str] = []
    for a in dxf_audits:
        for iss in a.issues:
            issues.append(f"[{a.plate_name}] {iss}")
    for iss in sim_3d.issues:
        issues.append(f"[Sim3D] {iss}")
    if pdf_overlay is not None:
        for iss in pdf_overlay.issues:
            issues.append(f"[OverlayPDF] {iss}")
    if cfg.exact_gabarit_lock and pdf_bom_len_err is not None and pdf_bom_len_err > 0.01:
        issues.append(f"Gabarit 3D {sim_3d.l_3d_mm:.3f} != STYKKLISTE {bom_row.length_mm:.3f}")

    passed = (
        all_dxf_clean
        and all_layer_zero
        and nc1_max_dim_err <= 0.01
        and nc1_max_hole_err <= 0.05
        and (not cfg.exact_gabarit_lock or pdf_bom_len_err is None or pdf_bom_len_err <= 0.01)
        and sim_3d.passed
        and (pdf_overlay is None or pdf_overlay.passed)
    )

    pdf_name = (
        bom.source_pdf.name
        if bom is not None
        else (pdf_analysis.pdf_name if pdf_analysis is not None else (pdf_overlay.pdf_name if pdf_overlay else None))
    )
    return PostDrawVerificationReport(
        piece_id=dstv.header.piece_id,
        plates_audited=len(dxf_audits),
        all_dxf_clean=all_dxf_clean,
        all_layer_zero=all_layer_zero,
        nc1_max_dim_err_mm=round(nc1_max_dim_err, 4),
        nc1_max_hole_err_mm=round(nc1_max_hole_err, 4),
        pdf_drawing_name=pdf_name,
        pdf_bom_length_err_mm=pdf_bom_len_err,
        pdf_dims_confirmed=pdf_confirmed,
        sim_3d=sim_3d,
        plate_audits=dxf_audits,
        passed=passed,
        issues=issues,
        pdf_text_source=(pdf_analysis.text_source if pdf_analysis else "none"),
        pdf_profile_callouts=(list(pdf_analysis.profile_callouts) if pdf_analysis else []),
        pdf_hole_callouts=(list(pdf_analysis.hole_callouts) if pdf_analysis else []),
        pdf_overlay=pdf_overlay,
    )


def find_oda_exe(custom_path: Optional[str] = None) -> Optional[str]:
    if custom_path and os.path.isfile(custom_path.strip()):
        return custom_path.strip()
    if PdfToCadConverter is not None:
        found = PdfToCadConverter._find_default_oda_exe()
        if found and os.path.isfile(found):
            return found
    return None


def convert_dxf_to_dwg(
    dxf_path: str | Path,
    dwg_path: str | Path,
    oda_path: Optional[str] = None,
    version: str = "ACAD2018",
) -> Path:
    exe = find_oda_exe(oda_path)
    if not exe:
        raise RuntimeError("ODAFileConverter.exe nu a fost găsit.")
    src = Path(dxf_path).resolve()
    dst = Path(dwg_path).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="steel_oda_") as tmp_root:
        in_dir = Path(tmp_root) / "in"
        out_dir = Path(tmp_root) / "out"
        in_dir.mkdir(parents=True, exist_ok=True)
        out_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, in_dir / src.name)

        cmd = [exe, str(in_dir), str(out_dir), version, "DWG", "0", "1", src.name]
        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        produced = out_dir / f"{src.stem}.dwg"
        if not produced.is_file():
            matches = list(out_dir.glob("*.dwg"))
            if matches:
                produced = matches[0]
            else:
                raise RuntimeError(f"ODA conversion failed (rc={proc.returncode}): {proc.stderr or proc.stdout}")
        shutil.copy2(produced, dst)
        return dst


# ─── PyMuPDF STYKKLISTE BOM Extractor ────────────────────────────────────────

@dataclass
class _WordBox:
    x0: float
    top: float
    x1: float
    bottom: float
    text: str

    @property
    def x_center(self) -> float:
        return 0.5 * (self.x0 + self.x1)


def _cluster_words_by_y(words: list[_WordBox], y_tol: float = 4.5) -> list[list[_WordBox]]:
    if not words:
        return []
    ordered = sorted(words, key=lambda w: (w.top, w.x0))
    rows: list[list[_WordBox]] = [[ordered[0]]]
    row_y = ordered[0].top
    for w in ordered[1:]:
        if abs(w.top - row_y) <= y_tol:
            rows[-1].append(w)
        else:
            rows[-1].sort(key=lambda item: item.x0)
            rows.append([w])
            row_y = w.top
    rows[-1].sort(key=lambda item: item.x0)
    return rows


def _parse_float_opt(text: str) -> Optional[float]:
    s = (text or "").strip().replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_int_opt(text: str, default: int = 1) -> int:
    val = _parse_float_opt(text)
    return int(round(val)) if val is not None else default


def _match_header_column(token_upper: str) -> Optional[str]:
    clean = token_upper.strip(":.")
    if clean in ("LITTRA", "LITTERA", "MARK"):
        return "littra"
    if clean in ("POS", "POS.", "POSITION", "PART"):
        return "pos_no"
    if clean in ("ANTALL", "ANTAL", "QTY", "QUANTITY", "STK"):
        return "qty"
    if clean in ("PROFIL", "PROFILE", "SECTION"):
        return "profile"
    if clean in ("MATERIAL", "MAT", "GRADE", "QUALITY"):
        return "material"
    if clean in ("LENGDE", "LENGTH", "LÄNGD", "LANGD"):
        return "length"
    if clean in ("VEKT/STK", "WEIGHT/PC", "UNIT", "VEKT", "WEIGHT"):
        return "weight_unit"
    if clean in ("TOT.", "TOT", "TOTAL"):
        return "weight_total"
    return None


def _ocr_pdf_page_image(page, dpi: int = 200) -> tuple[str, str]:
    """Run OCR (Tesseract if installed, else built-in Windows WinRT OcrEngine) on a PDF page image."""
    if page is None:
        return "", "none"

    pix = page.get_pixmap(dpi=dpi)
    with tempfile.TemporaryDirectory(prefix="steel_ocr_") as tmp_dir:
        img_path = Path(tmp_dir) / "page.png"
        pix.save(str(img_path))

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

        # Path 1: Tesseract CLI if installed
        tess = shutil.which("tesseract")
        if tess:
            try:
                proc = subprocess.run(
                    [tess, str(img_path), "stdout", "--psm", "6"],
                    capture_output=True,
                    text=True,
                    timeout=20,
                    startupinfo=startupinfo,
                    creationflags=creationflags,
                )
                txt = (proc.stdout or "").strip()
                if txt:
                    return txt, "ocr_tesseract"
            except Exception:
                pass

        # Path 2: Built-in Windows 10/11 WinRT OCR (zero external dependencies)
        if os.name == "nt":
            img_escaped = str(img_path).replace("'", "''")
            ps_script = f"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }})[0]
function Await($WinRtTask, $ResultType) {{
    $netTask = $asTaskGeneric.MakeGenericMethod($ResultType).Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}}
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync('{img_escaped}')) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($null -eq $engine) {{ $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language 'en-US')) }}
if ($null -ne $engine) {{
    $res = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $res.Text
}}
"""
            try:
                proc = subprocess.run(
                    ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=25,
                    startupinfo=startupinfo,
                    creationflags=creationflags,
                )
                txt = (proc.stdout or "").strip()
                if txt:
                    return txt, "ocr_winrt"
            except Exception:
                pass

    return "", "none"


def _parse_drawing_text_facts(full_text: str) -> tuple[set[float], list[str], list[tuple[int, float]]]:
    dims: set[float] = set()
    for m in re.finditer(r"(?<![A-Za-z0-9\-/])\d{2,5}(?:[.,]\d{1,2})?(?![A-Za-z0-9\-])", full_text):
        try:
            dims.add(float(m.group(0).replace(",", ".")))
        except ValueError:
            pass

    profiles: list[str] = []
    seen_prof: set[str] = set()
    for m in re.finditer(r"\b(?:HSQ|BOX|WQ)\d+(?:[.,]\d+)?\*\d+(?:[.,]\d+)?(?:-\d+(?:[.,]\d+)?\*(?:\d+(?:[.,]\d+)?|\*{2,3}))+", full_text, re.IGNORECASE):
        val = m.group(0).strip()
        if val not in seen_prof:
            seen_prof.add(val)
            profiles.append(val)

    hole_callouts: list[tuple[int, float]] = []
    seen_holes: set[tuple[int, float]] = set()
    for m in re.finditer(r"(\d+)\s*(?:stk|buc|pcs|x|×)\s*[Øø∅]\s*(\d+(?:[.,]\d+)?)", full_text, re.IGNORECASE):
        qty = int(m.group(1))
        dia = float(m.group(2).replace(",", "."))
        pair = (qty, dia)
        if pair not in seen_holes:
            seen_holes.add(pair)
            hole_callouts.append(pair)
    return dims, profiles, hole_callouts


def _extract_bom_from_ocr_text(ocr_text: str, source_pdf: Path) -> Optional[BomTable]:
    """Fallback BOM line parser for scanned PDFs when vector STYKKLISTE words are absent."""
    rows: list[BomRow] = []
    for raw_ln in ocr_text.splitlines():
        m = re.match(
            r"^\s*([A-Za-z]{1,6}[-_]?\d{2,6})\s+(\d{1,3})\s+((?:HSQ|BOX|WQ|PL|FL)\S+)\s+(S\d{3}\w*)\s+(\d{2,5}(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)",
            raw_ln.strip(),
            re.IGNORECASE,
        )
        if m:
            rows.append(
                BomRow(
                    pos_no=m.group(1),
                    littra="",
                    qty_per_assembly=int(m.group(2)),
                    profile=m.group(3),
                    material=m.group(4),
                    length_mm=_parse_float_opt(m.group(5)),
                    weight_unit_kg=_parse_float_opt(m.group(6)),
                    weight_total_kg=None,
                )
            )
    if not rows:
        return None
    return BomTable(
        assembly_id=source_pdf.stem,
        assembly_quantity=1,
        rows=rows,
        total_weight_kg=None,
        source_pdf=source_pdf,
    )


def extract_bom_from_pdf(pdf_path: str | Path, allow_ocr: bool = True) -> Optional[BomTable]:
    if pymupdf is None:
        return None
    path = Path(pdf_path)
    if not path.is_file():
        return None

    doc = pymupdf.open(str(path))
    try:
        for page in doc:
            raw_words = page.get_text("words") or []
            words = [
                _WordBox(float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4]).strip())
                for w in raw_words
                if str(w[4]).strip()
            ]
            if len(words) >= 5:
                vlines: set[float] = set()
                for drw in page.get_drawings():
                    for item in drw.get("items", []):
                        if item[0] == "l":
                            p1, p2 = item[1], item[2]
                            if abs(p1.x - p2.x) < 1.0 and abs(p1.y - p2.y) > 50.0:
                                vlines.add(float(p1.x))
                        elif item[0] == "re":
                            r = item[1]
                            if r.height > 50.0:
                                vlines.add(float(r.x0))
                                vlines.add(float(r.x1))

                bom = _parse_bom_from_words(words, float(page.rect.height), sorted(vlines), path)
                if bom is not None and bom.rows:
                    return bom
            elif allow_ocr:
                ocr_txt, _ = _ocr_pdf_page_image(page)
                if ocr_txt:
                    bom_ocr = _extract_bom_from_ocr_text(ocr_txt, path)
                    if bom_ocr is not None:
                        return bom_ocr
        return None
    finally:
        doc.close()


def extract_pdf_drawing_analysis(
    pdf_path: Optional[str | Path],
    allow_ocr: bool = True,
) -> Optional[PdfDrawingAnalysis]:
    """Extract dimensions, HSQ profile callouts, and hole callouts from PDF vector text or OCR fallback."""
    if pymupdf is None or not pdf_path:
        return None
    path = Path(pdf_path)
    if not path.is_file():
        return None

    doc = pymupdf.open(str(path))
    try:
        vec_chunks: list[str] = []
        vec_word_cnt = 0
        for page in doc:
            pw = page.get_text("words") or []
            vec_word_cnt += len(pw)
            pt = page.get_text("text") or ""
            if pt.strip():
                vec_chunks.append(pt)

        if vec_word_cnt >= 10:
            full_text = "\n".join(vec_chunks)
            dims, profs, holes = _parse_drawing_text_facts(full_text)
            for page in doc:
                for w in page.get_text("words") or []:
                    t = str(w[4]).strip()
                    if re.match(r"^\d{2,5}(?:[.,]\d{1,2})?$", t):
                        try:
                            dims.add(float(t.replace(",", ".")))
                        except ValueError:
                            pass
            return PdfDrawingAnalysis(
                pdf_name=path.name,
                text_source="vector",
                dimensions=dims,
                profile_callouts=profs,
                hole_callouts=holes,
                raw_char_count=len(full_text),
            )

        if allow_ocr and len(doc) > 0:
            ocr_texts: list[str] = []
            ocr_src = "none"
            for page in doc[:2]:
                txt, src = _ocr_pdf_page_image(page)
                if txt:
                    ocr_texts.append(txt)
                    ocr_src = src
            if ocr_texts:
                full_ocr = "\n".join(ocr_texts)
                dims, profs, holes = _parse_drawing_text_facts(full_ocr)
                return PdfDrawingAnalysis(
                    pdf_name=path.name,
                    text_source=ocr_src,
                    dimensions=dims,
                    profile_callouts=profs,
                    hole_callouts=holes,
                    raw_char_count=len(full_ocr),
                )

        return PdfDrawingAnalysis(pdf_name=path.name, text_source="none")
    finally:
        doc.close()


def extract_pdf_drawing_dimensions(pdf_path: Optional[str | Path]) -> set[float]:
    """Extract numeric dimensions from an A-drawing PDF (vector or OCR fallback)."""
    analysis = extract_pdf_drawing_analysis(pdf_path, allow_ocr=True)
    return set(analysis.dimensions) if analysis else set()


def _parse_bom_from_words(
    words: list[_WordBox],
    page_h: float,
    vlines: list[float],
    source_pdf: Path,
) -> Optional[BomTable]:
    if not words:
        return None

    banner_word: Optional[_WordBox] = None
    for w in words:
        up = w.text.upper()
        if any(kw in up for kw in BOM_BANNER_KEYWORDS):
            banner_word = w
            break
    if banner_word is None:
        return None

    banner_words = sorted(
        [
            w for w in words
            if abs(w.top - banner_word.top) <= 6.0
            and banner_word.x0 - 10.0 <= w.x0 <= banner_word.x0 + 620.0
        ],
        key=lambda w: w.x0,
    )
    assembly_id = ""
    assembly_qty = 1
    banner_tokens = [w.text for w in banner_words]
    for i, tok in enumerate(banner_tokens):
        up = tok.upper()
        if not assembly_id and up == "FOR" and i + 1 < len(banner_tokens):
            assembly_id = banner_tokens[i + 1].strip()
        if up == "ELEMENT" and i + 1 < len(banner_tokens):
            qv = _parse_float_opt(banner_tokens[i + 1])
            if qv is not None and qv > 0:
                assembly_qty = int(round(qv))

    # Restrict header search strictly within the STYKKLISTE table width (exclude right-hand title-block words)
    header_candidates = [
        w for w in words
        if banner_word.bottom - 2.0 <= w.top <= banner_word.bottom + 36.0
        and banner_word.x0 - 20.0 <= w.x0 <= banner_word.x0 + 610.0
    ]
    anchor_word: Optional[_WordBox] = None
    for w in header_candidates:
        ck = _match_header_column(w.text.upper())
        if ck in ("pos_no", "profile", "littra"):
            anchor_word = w
            break

    header_row: Optional[list[_WordBox]] = None
    if anchor_word is not None:
        header_row = sorted(
            [w for w in header_candidates if abs(w.top - anchor_word.top) <= 3.5],
            key=lambda w: w.x0,
        )
    else:
        for hr in _cluster_words_by_y(header_candidates, y_tol=4.5):
            joined = " ".join(x.text.upper() for x in hr)
            if ("POS" in joined or "PART" in joined) and ("PROFIL" in joined or "MAT" in joined):
                header_row = hr
                break
    if not header_row:
        return None

    col_boxes: dict[str, tuple[float, float]] = {}
    i = 0
    while i < len(header_row):
        w = header_row[i]
        up = w.text.upper()
        col_key = _match_header_column(up)
        if col_key == "pos_no" and i + 1 < len(header_row) and header_row[i + 1].text.upper().startswith("NO"):
            col_boxes[col_key] = (w.x0, header_row[i + 1].x1)
            i += 2
            continue
        if col_key == "weight_total" and i + 1 < len(header_row) and "VEKT" in header_row[i + 1].text.upper():
            col_boxes[col_key] = (w.x0, header_row[i + 1].x1)
            i += 2
            continue
        if col_key and col_key not in col_boxes:
            col_boxes[col_key] = (w.x0, w.x1)
        i += 1

    if "pos_no" not in col_boxes:
        return None

    ordered_cols = sorted(col_boxes.items(), key=lambda item: item[1][0])
    table_x0 = min(banner_word.x0, ordered_cols[0][1][0]) - 8.0
    last_x1 = ordered_cols[-1][1][1]
    border_candidates = [vx for vx in vlines if last_x1 - 2.0 <= vx <= last_x1 + 25.0]
    table_x1 = min(border_candidates) - 0.5 if border_candidates else (last_x1 + 6.0)

    col_bounds: dict[str, tuple[float, float]] = {}
    for idx_c, (cname, (cx0, _)) in enumerate(ordered_cols):
        lb = table_x0 if idx_c == 0 else (cx0 - (1.0 if cname == "material" else 2.5))
        if idx_c + 1 < len(ordered_cols):
            next_name, (nx0, _) = ordered_cols[idx_c + 1]
            rb = nx0 - (1.0 if next_name == "material" else 2.5)
        else:
            rb = table_x1
        col_bounds[cname] = (lb, rb)

    matched_header_words = [w for w in header_row if w.x0 <= last_x1 + 2.0]
    header_bottom = max(w.bottom for w in (matched_header_words or header_row))
    body_words = [
        w for w in words
        if w.top > header_bottom - 1.0 and table_x0 <= w.x_center <= table_x1 and w.top <= page_h
    ]

    def _assign_col(word: _WordBox) -> Optional[str]:
        xc = word.x_center
        for cname, (lb, rb) in col_bounds.items():
            if lb <= xc < rb:
                return cname
        for cname, (lb, rb) in col_bounds.items():
            if lb <= word.x0 < rb:
                return cname
        return None

    rows: list[BomRow] = []
    total_weight: Optional[float] = None
    for r_words in _cluster_words_by_y(body_words, y_tol=4.5):
        cells: dict[str, list[str]] = {}
        for w in r_words:
            cname = _assign_col(w)
            if cname:
                cells.setdefault(cname, []).append(w.text)

        pos_txt = " ".join(cells.get("pos_no", [])).strip()
        littra_txt = " ".join(cells.get("littra", [])).strip()
        qty_txt = " ".join(cells.get("qty", [])).strip()
        prof_txt = " ".join(cells.get("profile", [])).strip()
        mat_txt = " ".join(cells.get("material", [])).strip()
        len_txt = " ".join(cells.get("length", [])).strip()
        wunit_txt = " ".join(cells.get("weight_unit", [])).strip()
        wtot_txt = " ".join(cells.get("weight_total", [])).strip()

        row_all_upper = " ".join(w.text.upper() for w in r_words)
        if "TOTALT" in row_all_upper or (not pos_txt and "TOTAL" in row_all_upper):
            total_weight = _parse_float_opt(wtot_txt)
            if total_weight is None:
                for w in reversed(r_words):
                    v = _parse_float_opt(w.text)
                    if v is not None:
                        total_weight = v
                        break
            break

        if not pos_txt:
            continue

        rows.append(
            BomRow(
                pos_no=pos_txt,
                littra=littra_txt,
                qty_per_assembly=_parse_int_opt(qty_txt, 1),
                profile=prof_txt,
                material=mat_txt,
                length_mm=_parse_float_opt(len_txt),
                weight_unit_kg=_parse_float_opt(wunit_txt),
                weight_total_kg=_parse_float_opt(wtot_txt),
            )
        )

    if not rows:
        return None
    return BomTable(
        assembly_id=assembly_id or source_pdf.stem,
        assembly_quantity=assembly_qty,
        rows=rows,
        total_weight_kg=total_weight,
        source_pdf=source_pdf,
    )


def parse_plate_profile(profile_str: str) -> Optional[tuple[float, float]]:
    m = _PLATE_PROFILE_RE.match((profile_str or "").strip())
    if not m:
        return None
    return float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", "."))


def _is_w_drawing_pdf_name(stem_lower: str) -> bool:
    return bool(
        "w-drawing" in stem_lower
        or "w_drawing" in stem_lower
        or re.search(r"[-_\s]w$", stem_lower)
    )


def find_assembly_a_drawing(
    folder: Path,
    target_piece_keys: Optional[set[str]] = None,
) -> Optional[Path]:
    candidate_dirs = [folder]
    if folder.parent != folder:
        candidate_dirs.append(folder.parent)

    for cand_idx, cand_dir in enumerate(candidate_dirs):
        if not cand_dir.is_dir():
            continue
        pdfs = sorted(cand_dir.glob("*.pdf"))
        if not pdfs:
            continue

        # Pass 1: explicit A-drawing / assembly filenames
        preferred = [
            p for p in pdfs
            if not _is_w_drawing_pdf_name(p.stem.lower())
            and (
                "a-drawing" in p.stem.lower()
                or "a_drawing" in p.stem.lower()
                or "assembly" in p.stem.lower()
            )
        ]
        for p in preferred:
            if pymupdf is None:
                return p
            bom = extract_bom_from_pdf(p)
            if bom is not None and bom.rows:
                return p
        if preferred and pymupdf is None:
            return preferred[0]

        # Pass 2: fallback scan of any non-W PDF containing a STYKKLISTE table
        if pymupdf is not None:
            others = [
                p for p in pdfs
                if p not in preferred and not _is_w_drawing_pdf_name(p.stem.lower())
            ]
            max_scan = 20 if cand_idx == 0 else 12
            best_pdf: Optional[Path] = None
            best_score = (-1, -1)
            for p in others[:max_scan]:
                try:
                    bom = extract_bom_from_pdf(p)
                except Exception:
                    bom = None
                if bom is not None and bom.rows:
                    matched = 0
                    if target_piece_keys:
                        row_keys = {clean_piece_name(r.pos_no, lowercase=True) for r in bom.rows}
                        matched = len(target_piece_keys & row_keys)
                    score = (matched, len(bom.rows))
                    if score > best_score:
                        best_score = score
                        best_pdf = p
                    if target_piece_keys and matched > 0:
                        return p
            if best_pdf is not None:
                return best_pdf

    return None


def discover_w_drawings(folder: Path) -> dict[str, Path]:
    w_map: dict[str, Path] = {}
    for p in sorted(folder.glob("*.pdf")):
        low = p.stem.lower()
        if "w-drawing" in low or "w_drawing" in low or re.search(r"[-_\s]w$", low):
            prefix = re.split(r"[-_\s]+w(?:[-_\s]*drawing)?", p.stem, maxsplit=1, flags=re.IGNORECASE)[0]
            key = clean_piece_name(prefix, lowercase=True)
            if key:
                w_map[key] = p
    return w_map


# ─── Preview & Unified Pipeline (GUI & CLI) ──────────────────────────────────

def _export_plate_cad(
    plate: ExtractedPlate,
    out_dir: Path,
    export_format: str,
    use_nesting_layers: bool,
    include_marking: bool,
    organize_by_thickness: bool,
    oda_path: Optional[str],
    log_cb: Callable[[str], None],
) -> tuple[list[Path], str, DrawnDxfAudit, Ring]:
    target_dir = (out_dir / plate.folder) if organize_by_thickness else out_dir
    target_dir.mkdir(parents=True, exist_ok=True)
    dxf_path = target_dir / f"{plate.name}.dxf"
    write_plate_dxf(
        plate,
        dxf_path,
        use_nesting_layers=use_nesting_layers,
        include_marking=include_marking,
    )
    dxf_audit, re_read_ring = inspect_drawn_dxf(
        dxf_path, plate, expect_layer_zero=(not use_nesting_layers)
    )

    fmt = (export_format or "dwg_fallback").lower()
    if fmt == "dxf":
        rel = f"{plate.folder}/{plate.name}.dxf" if organize_by_thickness else f"{plate.name}.dxf"
        return [dxf_path], rel, dxf_audit, re_read_ring

    dwg_path = target_dir / f"{plate.name}.dwg"
    exe = find_oda_exe(oda_path)
    if exe:
        try:
            convert_dxf_to_dwg(dxf_path, dwg_path, oda_path=exe)
            if fmt == "both":
                rel = f"{plate.folder}/{plate.name}.dwg" if organize_by_thickness else f"{plate.name}.dwg"
                return [dxf_path, dwg_path], rel, dxf_audit, re_read_ring
            try:
                dxf_path.unlink()
            except Exception:
                pass
            rel = f"{plate.folder}/{plate.name}.dwg" if organize_by_thickness else f"{plate.name}.dwg"
            return [dwg_path], rel, dxf_audit, re_read_ring
        except Exception as exc:
            log_cb(f"  ⚠ Eroare conversie ODA pentru {plate.name}: {exc} (păstrat DXF)")

    if fmt == "dwg_direct":
        raise RuntimeError(f"Conversia DWG direct a eșuat pentru {plate.name} (ODA indisponibil).")

    rel = f"{plate.folder}/{plate.name}.dxf" if organize_by_thickness else f"{plate.name}.dxf"
    return [dxf_path], rel, dxf_audit, re_read_ring


_EXCLUDED_BATCH_DIR_NAMES = {
    "cheson_output",
    "test",
    "tests",
    "backup",
    "old",
    "tmp",
    "temp",
    "w_drawing",
}
_THICKNESS_DIR_RE = re.compile(r"^t\d+(?:[.,]\d+)?$", re.IGNORECASE)


def _is_excluded_batch_dir(dir_name: str) -> bool:
    norm = dir_name.strip().lower()
    return (
        norm in _EXCLUDED_BATCH_DIR_NAMES
        or norm.startswith(".")
        or bool(_THICKNESS_DIR_RE.match(norm))
    )


def discover_cheson_batch_folders(root: Path, recursive: bool = False) -> list[Path]:
    """Discover isolated cheson folders under `root` containing direct .nc1 files.

    - Walks non-excluded subdirectories under `root` (skipping `test`, `T5..T25`,
      `Cheson_Output`, etc.). If any cheson subdirectories are found, returns them
      even if `root` also contains loose .nc1 files alongside the subfolders.
    - Otherwise, if `root` itself directly contains .nc1 files, returns `[root]`.
    """
    root = Path(root).resolve()
    if not root.is_dir():
        return []

    discovered_subdirs: list[Path] = []
    try:
        top_children = sorted(root.iterdir(), key=lambda p: p.name.lower())
    except OSError:
        return []

    stack: list[Path] = [
        c for c in top_children if c.is_dir() and not _is_excluded_batch_dir(c.name)
    ]
    while stack:
        current = stack.pop(0)
        try:
            children = sorted(current.iterdir(), key=lambda p: p.name.lower())
        except OSError:
            continue

        direct_nc1 = any(c.is_file() and c.suffix.lower() == ".nc1" for c in children)
        if direct_nc1:
            discovered_subdirs.append(current)

        subdirs = [c for c in children if c.is_dir() and not _is_excluded_batch_dir(c.name)]
        stack = subdirs + stack

    if discovered_subdirs:
        return discovered_subdirs

    has_direct_nc1 = any(p.is_file() and p.suffix.lower() == ".nc1" for p in top_children)
    if has_direct_nc1:
        return [root]

    return []


def _collect_inputs(
    source_path: str | Path,
    source_mode: str,
    scope_mode: str,
    recursive: bool,
) -> tuple[Path, Path, list[Path], Optional[Path], Optional[BomTable], dict[str, Path]]:
    src = Path(source_path).resolve()
    if source_mode == "single_file" or src.is_file():
        if not src.is_file() or src.suffix.lower() != ".nc1":
            raise FileNotFoundError(f"Fișierul .nc1 nu este valid: {src}")
        nc1_files = [src]
        search_root = src.parent
    else:
        if not src.is_dir():
            raise FileNotFoundError(f"Folderul sursă nu există: {src}")
        search_root = src
        nc1_files = sorted(src.rglob("*.nc1") if recursive else src.glob("*.nc1"))

    if not nc1_files and scope_mode == "cheson_only":
        raise FileNotFoundError(f"Nu s-au găsit fișiere .nc1 în {src}")

    # Always read A-drawing / STYKKLISTE BOM silently if present so exact LENGDE is available
    target_keys = {clean_piece_name(p.stem, lowercase=True) for p in nc1_files}
    a_pdf = find_assembly_a_drawing(search_root, target_piece_keys=target_keys) if search_root.is_dir() else None
    bom = extract_bom_from_pdf(a_pdf) if a_pdf is not None else None
    w_map = discover_w_drawings(search_root) if (search_root.is_dir() and scope_mode != "cheson_only") else {}
    return src, search_root, nc1_files, a_pdf, bom, w_map


def extract_pdf_noted_dimensions(pdf_path: Optional[str | Path]) -> set[float]:
    """Extract clean drawing dimension callouts from TOP/FRONT/SECTION (excluding dates, scales, Ø, BOM)."""
    if pymupdf is None or not pdf_path:
        return set()
    path = Path(pdf_path)
    if not path.is_file():
        return set()

    dims: set[float] = set()
    doc = pymupdf.open(str(path))
    try:
        for page in doc:
            words = page.get_text("words") or []
            # Find STYKKLISTE titleblock region to avoid BOM weights/dates/scales
            bom_top = page.rect.height
            bom_x0 = page.rect.width
            for w in words:
                up = str(w[4]).strip().upper()
                if any(kw in up for kw in BOM_BANNER_KEYWORDS):
                    bom_top = min(bom_top, float(w[1]) - 15.0)
                    bom_x0 = min(bom_x0, float(w[0]) - 20.0)
            for idx, w in enumerate(words):
                x0, y0, x1, y1, raw = float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4]).strip()
                if x0 >= bom_x0 and y0 >= bom_top:
                    continue
                if re.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$", raw):
                    continue
                if raw.startswith(("Ø", "ø", "∅", "R", "M", "S", "T", "1:")):
                    continue
                if re.match(r"^\d{2,5}(?:[.,]\d{1,2})?$", raw):
                    # Skip scale denominator right after "1:" or diameter value right after Ø
                    prev = str(words[idx - 1][4]).strip() if idx > 0 else ""
                    if prev in ("Ø", "ø", "∅", "stk", "×", "x", "1:"):
                        continue
                    try:
                        val = float(raw.replace(",", "."))
                        if val >= 20.0:
                            dims.add(val)
                    except ValueError:
                        pass
    finally:
        doc.close()
    return dims


def _clone_plate_with_geometry(
    p: ExtractedPlate,
    exterior: Ring,
    holes: list[DSTVHole],
    inners: list[Ring],
    markings: list[DSTVMarking],
    source_bounds: tuple[float, float, float, float],
) -> ExtractedPlate:
    return ExtractedPlate(
        name=p.name,
        folder=p.folder,
        role=p.role,
        face=p.face,
        thickness=p.thickness,
        piece_id=p.piece_id,
        piece_clean=p.piece_clean,
        material=p.material,
        quantity=p.quantity,
        exterior=exterior,
        holes=holes,
        inner_contours=inners,
        markings=markings,
        source_bounds=source_bounds,
        removed_residuals=p.removed_residuals,
        qty_per_cheson=p.qty_per_cheson,
        role_code=p.role_code,
        merged_sources=list(p.merged_sources),
    )


def _snap_plate_rings_to_noted_dims(
    dstv: DSTVFile,
    plates: list[ExtractedPlate],
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
    detected_type: str,
    talpa_side: str,
    l_nc1_raw: float,
    l_gabarit: float,
    l_source_label: str,
    noted_dims: set[float],
) -> tuple[list[ExtractedPlate], list[str]]:
    """Snap near-integer plate lengths (<=0.55 mm) to noted PDF drawing dims only when 3D seams remain intact."""
    if not cfg.exact_gabarit_lock or not noted_dims or not plates:
        return plates, []

    notes: list[str] = []
    cur_plates = list(plates)
    for idx, p in enumerate(cur_plates):
        b0, b1, b2, b3 = p.bounds
        pL = b2 - b0
        if pL < 200.0:
            continue
        # Only snap if NOT already within 0.001 mm of a noted PDF dim, and find unique match within 0.55 mm
        cands = [d for d in noted_dims if abs(d - pL) <= 0.55]
        if not cands:
            continue
        best_d = min(cands, key=lambda d: abs(d - pL))
        delta = best_d - pL
        if abs(delta) <= 0.001:
            continue

        sx0, sy0, sx1, sy1 = p.source_bounds
        # Determine which end can move cleanly without exceeding [0, l_gabarit]
        move_right = (sx1 + delta <= l_gabarit + 1e-4) and (abs(sx1 - l_gabarit) > 0.6 or abs(delta) <= 0.01)
        move_left = (not move_right) and (sx0 - delta >= -1e-4) and (sx0 > 0.6)
        if not move_right and not move_left:
            continue

        mir_x = cfg.mirror_web_h_x if p.face == "h" else False
        if move_right:
            new_ext = clean_ring(
                [(x + delta if abs(x - pL) <= 0.15 else x, y) for x, y in p.exterior],
                min_seg_len=cfg.min_segment_length_mm,
                ccw=True,
            )
            new_holes = [
                DSTVHole(
                    face=h.face,
                    x=(h.x + delta if h.x > 0.5 * pL else h.x),
                    y=h.y,
                    diameter=h.diameter,
                    depth=h.depth,
                    hole_type=h.hole_type,
                    slot_length=h.slot_length,
                    slot_width=h.slot_width,
                    slot_angle=h.slot_angle,
                    x_flag=h.x_flag,
                    y_flag=h.y_flag,
                )
                for h in p.holes
            ]
            new_sb = (sx0, sy0, sx1 + delta, sy1)
        else:
            # Lengthen/shorten left end while keeping normalized origin at (0, 0)
            new_ext = clean_ring(
                [(x + delta if x > 0.15 else x, y) for x, y in p.exterior],
                min_seg_len=cfg.min_segment_length_mm,
                ccw=True,
            )
            new_holes = [
                DSTVHole(
                    face=h.face,
                    x=(h.x + (delta if mir_x else delta)),
                    y=h.y,
                    diameter=h.diameter,
                    depth=h.depth,
                    hole_type=h.hole_type,
                    slot_length=h.slot_length,
                    slot_width=h.slot_width,
                    slot_angle=h.slot_angle,
                    x_flag=h.x_flag,
                    y_flag=h.y_flag,
                )
                for h in p.holes
            ]
            new_sb = (sx0 - delta, sy0, sx1, sy1)

        cand_p = _clone_plate_with_geometry(p, new_ext, new_holes, p.inner_contours, p.markings, new_sb)
        trial_plates = list(cur_plates)
        trial_plates[idx] = cand_p
        trial_ver = verify_cheson_decomposition(
            dstv=dstv,
            plates=trial_plates,
            l_nc1_raw=l_nc1_raw,
            l_gabarit_target=l_gabarit,
            l_source_label=l_source_label,
            l_assembled=l_gabarit,
            detected_type=detected_type,
            talpa_side=talpa_side,
            spec=spec,
            cfg=cfg,
        )
        if trial_ver.passed and trial_ver.flange_web_contact_error_mm <= 0.05:
            cur_plates = trial_plates
            notes.append(f"{p.name}: L {pL:.3f} -> {best_d:.3f} mm (cotă desen)")

    return cur_plates, notes


def _score_closed_loop_candidate(ver: VerificationReport, post_ver: PostDrawVerificationReport) -> tuple[int, int, float, float]:
    geom_ok = 1 if (ver.passed and post_ver.all_dxf_clean and post_ver.all_layer_zero and post_ver.sim_3d.passed) else 0
    ov_ok = 1 if (post_ver.pdf_overlay is None or post_ver.pdf_overlay.passed) else 0
    ov = post_ver.pdf_overlay.overall_overlay_pct if post_ver.pdf_overlay else 100.0
    return (1 if (geom_ok and ov_ok) else 0, geom_ok, ov, -ver.max_gabarit_error_mm)


def _extract_and_verify_cheson_closed_loop(
    dstv: DSTVFile,
    base_cfg: ChesonConfig,
    bom: Optional[BomTable],
    bom_row: Optional[BomRow],
    pdf_analysis: Optional[PdfDrawingAnalysis],
    noted_dims: set[float],
    a_pdf: Optional[Path],
    cached_pdf_views: Optional[dict[str, Any]],
    use_nesting_layers: bool,
    tmp_dir: Path,
) -> tuple[
    ChesonConfig,
    list[ExtractedPlate],
    VerificationReport,
    list[ExtractedPlate],
    list[tuple[int, str]],
    list[DrawnDxfAudit],
    dict[str, Ring],
    PostDrawVerificationReport,
]:
    """Run extraction -> DXF write/read -> 3D sim -> PDF overlay, auto-refining if <99.9% or check fails."""
    spec_base = _align_spec_flanges_to_faces(dstv, parse_compound_profile(dstv.header.profile_str, dstv.header), base_cfg)
    pdf_dims = pdf_analysis.dimensions if pdf_analysis else set()

    def _clone_cfg(**kw) -> ChesonConfig:
        d = asdict(base_cfg)
        d.update(kw)
        return ChesonConfig(**d)

    # Candidate strategies in priority order: (label, cfg_override, allow_noted_snap)
    strategies: list[tuple[str, ChesonConfig, bool]] = [
        ("Standard + Cote PDF", base_cfg, True),
        ("Fără ajustare cote intermediare", base_cfg, False),
        ("Păstrează segmente marginale (strip_residual=False)", _clone_cfg(strip_residual_fragments=False), False),
        ("Plan oblic față exterioară (skew=outer)", _clone_cfg(web_skew_plane="outer"), False),
        ("Plan oblic axă (skew=centerline)", _clone_cfg(web_skew_plane="centerline"), False),
    ]
    if base_cfg.cheson_type == "auto":
        det_t, _ = detect_cheson_type(dstv, spec_base, base_cfg)
        if det_t == "cheson_fara_talpa":
            strategies.append(
                (
                    "Capace peste inimi (over_webs)",
                    _clone_cfg(flange_top_mode="over_webs", flange_bot_mode="over_webs", web_skew_plane="outer"),
                    False,
                )
            )
    else:
        auto_recovery_cfg = ChesonConfig.from_preset(
            "auto",
            exact_gabarit_lock=base_cfg.exact_gabarit_lock,
            target_length_mm=base_cfg.target_length_mm,
            override_web_height_mm=base_cfg.override_web_height_mm,
            strip_residual_fragments=base_cfg.strip_residual_fragments,
            deduplicate_identical=base_cfg.deduplicate_identical,
        )
        strategies.append(("Auto-Detect Tip Cheson (corecție tip)", auto_recovery_cfg, True))

    best_bundle = None
    best_score = (-1, -1, -1.0, -999.0)
    attempt_notes: list[str] = []
    iters_done = 0

    for idx_s, (label, cand_cfg, allow_snap) in enumerate(strategies):
        # Skip duplicate strategy #1 if no noted dims would have snapped anyway
        if idx_s == 1 and best_bundle is not None and not best_bundle[8]:
            continue
        iters_done += 1
        cand_raw, cand_ver = decompose_nc1(dstv, cand_cfg, return_verification=True)
        snapped_notes: list[str] = []
        if allow_snap and noted_dims:
            cand_raw_s, snapped_notes = _snap_plate_rings_to_noted_dims(
                dstv=dstv,
                plates=cand_raw,
                spec=spec_base,
                cfg=cand_cfg,
                detected_type=cand_ver.detected_type,
                talpa_side=detect_cheson_type(dstv, spec_base, cand_cfg)[1],
                l_nc1_raw=cand_ver.l_nc1_raw_mm,
                l_gabarit=cand_ver.l_gabarit_target_mm,
                l_source_label=cand_ver.l_source_label,
                noted_dims=noted_dims,
            )
            if snapped_notes:
                cand_raw = cand_raw_s

        cand_unique, cand_map = deduplicate_cheson_plates(cand_raw, cand_cfg)
        cand_audits: list[DrawnDxfAudit] = []
        cand_rings: dict[str, Ring] = {}
        sub_tmp = tmp_dir / f"try_{idx_s}"
        sub_tmp.mkdir(parents=True, exist_ok=True)
        for p in cand_unique:
            tdxf = sub_tmp / f"{p.name}.dxf"
            write_plate_dxf(p, tdxf, use_nesting_layers=use_nesting_layers, include_marking=False)
            aud, rr = inspect_drawn_dxf(tdxf, p, expect_layer_zero=(not use_nesting_layers))
            cand_audits.append(aud)
            cand_rings[p.name] = rr

        exp_auds, exp_rings = _expand_audits_for_3d_sim(
            cand_raw, cand_unique, cand_map, cand_audits, cand_rings
        )
        cand_post = verify_post_draw_and_3d(
            dstv=dstv,
            cfg=cand_cfg,
            plates=cand_raw,
            dxf_audits=exp_auds,
            re_read_rings=exp_rings,
            bom=bom,
            bom_row=bom_row,
            pdf_dims=pdf_dims,
            pdf_analysis=pdf_analysis,
            a_pdf_path=a_pdf,
            cached_pdf_views=cached_pdf_views,
        )
        sc = _score_closed_loop_candidate(cand_ver, cand_post)
        ov_s = f"{cand_post.pdf_overlay.overall_overlay_pct:.2f}%" if cand_post.pdf_overlay else "n/a"
        attempt_notes.append(f"#{idx_s + 1} {label}: pass={cand_ver.passed and cand_post.passed}, overlay={ov_s}")

        if sc > best_score:
            best_score = sc
            best_bundle = (
                cand_cfg, cand_raw, cand_ver, cand_unique, cand_map,
                cand_audits, cand_rings, cand_post, snapped_notes,
            )
        if sc[0] == 1 and (cand_post.pdf_overlay is None or cand_post.pdf_overlay.overall_overlay_pct >= 99.9):
            break

    win_cfg, win_raw, win_ver, win_unique, win_map, win_auds, win_rings, win_post, win_snaps = best_bundle
    win_post.refinement_iterations = iters_done
    win_post.refinement_notes = attempt_notes if iters_done > 1 else []
    win_post.snapped_noted_dims = win_snaps
    return win_cfg, win_raw, win_ver, win_unique, win_map, win_auds, win_rings, win_post


def preview_cheson_pipeline(
    source_path: str | Path,
    source_mode: str = "folder",
    scope_mode: str = "cheson_only",
    preset: str = "auto",
    exact_gabarit_lock: bool = True,
    override_length_mm: Optional[float] = None,
    override_web_height_mm: Optional[float] = None,
    recursive: bool = False,
    strip_residual_fragments: bool = True,
    use_nesting_layers: bool = False,
    deduplicate_identical: bool = True,
    cheson_report_xlsx: Optional[str | Path] = None,
) -> dict[str, Any]:
    """Dry-run inspection: detects cheson type, STYKKLISTE/Excel qty, dedup T/C/L, DXF audit & 3D sim."""
    src_path = Path(source_path).resolve()
    if source_mode != "single_file" and src_path.is_dir():
        batch_folders = discover_cheson_batch_folders(src_path, recursive=recursive)
        is_multi = len(batch_folders) > 1 or (len(batch_folders) == 1 and batch_folders[0] != src_path)
        if is_multi:
            resolved_xlsx = find_cheson_report_xlsx(src_path, cheson_report_xlsx)
            all_items: list[dict[str, Any]] = []
            all_passed = True
            total_bom_rows = 0
            total_w_count = 0
            a_drawings_count = 0
            folder_names: list[str] = []
            for sub_dir in batch_folders:
                rel_name = str(sub_dir.relative_to(src_path)).replace("\\", "/")
                folder_names.append(rel_name)
                sub_prev = preview_cheson_pipeline(
                    source_path=sub_dir,
                    source_mode="folder",
                    scope_mode=scope_mode,
                    preset=preset,
                    exact_gabarit_lock=exact_gabarit_lock,
                    override_length_mm=override_length_mm,
                    override_web_height_mm=override_web_height_mm,
                    recursive=False,
                    strip_residual_fragments=strip_residual_fragments,
                    use_nesting_layers=use_nesting_layers,
                    deduplicate_identical=deduplicate_identical,
                    cheson_report_xlsx=resolved_xlsx or cheson_report_xlsx,
                )
                if sub_prev.get("a_drawing"):
                    a_drawings_count += 1
                total_bom_rows += int(sub_prev.get("bom_rows_count", 0))
                total_w_count += int(sub_prev.get("w_drawings_count", 0))
                if not sub_prev.get("all_passed", True):
                    all_passed = False
                for item in sub_prev.get("nc1_items", []):
                    item["folder_name"] = rel_name
                    item["a_drawing"] = sub_prev.get("a_drawing")
                    item["bom_assembly_id"] = sub_prev.get("bom_assembly_id")
                    all_items.append(item)

            return {
                "is_multi_folder_batch": True,
                "folders_count": len(batch_folders),
                "folders": folder_names,
                "a_drawing": f"{a_drawings_count}/{len(batch_folders)} foldere cu A-drawing",
                "cheson_report_xlsx": resolved_xlsx.name if resolved_xlsx else None,
                "bom_assembly_id": None,
                "bom_rows_count": total_bom_rows,
                "w_drawings_count": total_w_count,
                "all_passed": all_passed,
                "nc1_items": all_items,
            }

    _, search_root, nc1_files, a_pdf, bom, w_map = _collect_inputs(
        source_path, source_mode, scope_mode, recursive
    )
    resolved_xlsx = find_cheson_report_xlsx(search_root, cheson_report_xlsx)
    excel_qty_map = load_cheson_quantities_from_excel(resolved_xlsx)
    pdf_analysis = extract_pdf_drawing_analysis(a_pdf)
    pdf_dims = pdf_analysis.dimensions if pdf_analysis else set()
    noted_dims = extract_pdf_noted_dimensions(a_pdf)
    cached_pdf_views = _extract_pdf_orthographic_views(a_pdf)
    bom_by_key: dict[str, BomRow] = {}
    if bom:
        for r in bom.rows:
            bom_by_key[clean_piece_name(r.pos_no, lowercase=True)] = r

    items_preview: list[dict[str, Any]] = []
    all_passed = True

    with tempfile.TemporaryDirectory(prefix="steel_prev_") as tmp_root:
        tmp_dir = Path(tmp_root)
        for nc1_path in nc1_files:
            dstv = parse_dstv(nc1_path)
            piece_key = clean_piece_name(dstv.header.piece_id or nc1_path.stem, lowercase=True)
            bom_row = bom_by_key.get(piece_key)
            target_len = override_length_mm if (override_length_mm and override_length_mm > 0) else (
                bom_row.length_mm if (bom_row and bom_row.length_mm) else None
            )
            base_cfg = ChesonConfig.from_preset(
                preset,
                exact_gabarit_lock=exact_gabarit_lock,
                target_length_mm=target_len,
                override_web_height_mm=override_web_height_mm,
                strip_residual_fragments=strip_residual_fragments,
                deduplicate_identical=deduplicate_identical,
            )
            piece_tmp = tmp_dir / piece_key
            piece_tmp.mkdir(parents=True, exist_ok=True)
            _, raw_plates, ver, unique_plates, _, _, _, post_ver = _extract_and_verify_cheson_closed_loop(
                dstv=dstv,
                base_cfg=base_cfg,
                bom=bom,
                bom_row=bom_row,
                pdf_analysis=pdf_analysis,
                noted_dims=noted_dims,
                a_pdf=a_pdf,
                cached_pdf_views=cached_pdf_views,
                use_nesting_layers=use_nesting_layers,
                tmp_dir=piece_tmp,
            )
            if not (ver.passed and post_ver.passed):
                all_passed = False

            if piece_key in excel_qty_map and resolved_xlsx is not None:
                cheson_qty = excel_qty_map[piece_key]
                qty_source = resolved_xlsx.name
            elif bom_row is not None:
                cheson_qty = max(1, bom_row.qty_per_assembly) * max(1, bom.assembly_quantity if bom else 1)
                qty_source = "STYKKLISTE"
            else:
                cheson_qty = max(1, dstv.header.quantity)
                qty_source = "NC1"

            plate_dicts: list[dict[str, Any]] = []
            for p in unique_plates:
                tot_q = p.qty_per_cheson * cheson_qty
                plate_dicts.append(
                    {
                        "name": p.name,
                        "folder": p.folder,
                        "role": p.role,
                        "role_code": p.role_code,
                        "role_label": _ROLE_LABELS_RO.get(p.role, p.role),
                        "thickness": p.thickness,
                        "length_mm": round(p.bounds[2] - p.bounds[0], 4),
                        "width_mm": round(p.bounds[3] - p.bounds[1], 4),
                        "removed_residuals": p.removed_residuals,
                        "qty_per_cheson": p.qty_per_cheson,
                        "cheson_qty": cheson_qty,
                        "total_qty": tot_q,
                        "merged_sources": list(p.merged_sources),
                    }
                )
            tcl_summary = format_cheson_tcl_summary(plate_dicts)

            items_preview.append(
                {
                    "file": nc1_path.name,
                    "piece_id": dstv.header.piece_id or nc1_path.stem,
                    "profile": dstv.header.profile_str,
                    "cheson_qty": cheson_qty,
                    "qty_source": qty_source,
                    "raw_plates_count": len(raw_plates),
                    "unique_plates_count": len(unique_plates),
                    "tcl_summary": tcl_summary,
                    "verification": asdict(ver),
                    "post_draw_audit": asdict(post_ver),
                    "plates": plate_dicts,
                }
            )

    return {
        "is_multi_folder_batch": False,
        "folders_count": 1,
        "a_drawing": a_pdf.name if a_pdf else None,
        "cheson_report_xlsx": resolved_xlsx.name if resolved_xlsx else None,
        "bom_assembly_id": bom.assembly_id if bom else None,
        "bom_rows_count": len(bom.rows) if bom else 0,
        "w_drawings_count": len(w_map),
        "all_passed": all_passed,
        "nc1_items": items_preview,
    }


def _run_cheson_single_folder(
    source_path: str | Path,
    dest_dir: str | Path,
    source_mode: str = "folder",
    scope_mode: str = "cheson_only",
    preset: str = "auto",
    export_format: str = "dwg_fallback",
    use_nesting_layers: bool = False,
    include_marking: bool = False,
    organize_by_thickness: bool = True,
    recursive: bool = False,
    generate_reports: bool = True,
    exact_gabarit_lock: bool = True,
    override_length_mm: Optional[float] = None,
    override_web_height_mm: Optional[float] = None,
    oda_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    strip_residual_fragments: bool = True,
    deduplicate_identical: bool = True,
    cheson_report_xlsx: Optional[str | Path] = None,
    _write_excel_report: bool = True,
) -> dict[str, Any]:
    """Run box-beam (.nc1) plate extraction on a single file or isolated cheson folder."""
    def _log(msg: str) -> None:
        if log_callback:
            log_callback(msg)

    def _prog(val: float) -> None:
        if progress_callback:
            progress_callback(max(0.0, min(1.0, val)))

    src, search_root, nc1_files, a_pdf, bom, w_map = _collect_inputs(
        source_path, source_mode, scope_mode, recursive
    )
    resolved_xlsx = find_cheson_report_xlsx(search_root, cheson_report_xlsx)
    excel_qty_map = load_cheson_quantities_from_excel(resolved_xlsx)
    pdf_analysis = extract_pdf_drawing_analysis(a_pdf)
    pdf_dims = pdf_analysis.dimensions if pdf_analysis else set()
    noted_dims = extract_pdf_noted_dimensions(a_pdf)
    cached_pdf_views = _extract_pdf_orthographic_views(a_pdf)
    out_root = Path(dest_dir).resolve()
    out_root.mkdir(parents=True, exist_ok=True)
    effective_oda = find_oda_exe(oda_path)

    layer_label = "CUT_OUTER/INNER/MARKING" if use_nesting_layers else "Layer 0 (implicit, fără layere extra)"
    _log(
        f"🚀 Pornire Cheson / HSQ Extractor (Mod: {scope_mode}, Preset: {preset}, "
        f"GabaritFix={exact_gabarit_lock}, Dedup={deduplicate_identical}, Layer={layer_label})"
    )
    _log(f"📁 Sursă: {src}")
    _log(f"📁 Destinație: {out_root}")
    if a_pdf is not None and bom is not None:
        _log(f"📋 STYKKLISTE ({a_pdf.name}): {bom.assembly_id} (qty={bom.assembly_quantity}, {len(bom.rows)} repere, {len(pdf_dims)} cote PDF)")
    if resolved_xlsx is not None and _write_excel_report:
        _log(f"📊 Raport cantități chesoane: {resolved_xlsx.name} ({len(excel_qty_map)} repere)")
    if export_format != "dxf":
        _log(f"🔧 ODA Converter: {effective_oda if effective_oda else 'indisponibil (fallback automat pe DXF R2018)'}")

    bom_by_key: dict[str, BomRow] = {}
    if bom:
        for r in bom.rows:
            bom_by_key[clean_piece_name(r.pos_no, lowercase=True)] = r

    assembly_qty = bom.assembly_quantity if bom else 1
    assembly_id = bom.assembly_id if bom else search_root.name

    entries: list[ManifestEntry] = []
    cheson_summaries: list[dict[str, Any]] = []
    verifications: list[dict[str, Any]] = []
    post_draw_audits: list[dict[str, Any]] = []
    written_files: list[Path] = []
    processed_keys: set[str] = set()
    nc1_plate_count = 0
    w_count = 0
    bom_plate_count = 0
    skipped_count = 0
    errors: list[str] = []

    total_steps = max(
        1,
        len(nc1_files)
        + (len(w_map) if scope_mode != "cheson_only" else 0)
        + (len(bom.rows) if (bom and scope_mode == "full_assembly") else 0),
    )
    done_steps = 0

    for nc1_path in nc1_files:
        try:
            dstv = parse_dstv(nc1_path)
            piece_key = clean_piece_name(dstv.header.piece_id or nc1_path.stem, lowercase=True)
            processed_keys.add(piece_key)
            bom_row = bom_by_key.get(piece_key)
            target_len = override_length_mm if (override_length_mm and override_length_mm > 0) else (
                bom_row.length_mm if (bom_row and bom_row.length_mm) else None
            )
            base_cfg = ChesonConfig.from_preset(
                preset,
                exact_gabarit_lock=exact_gabarit_lock,
                target_length_mm=target_len,
                override_web_height_mm=override_web_height_mm,
                strip_residual_fragments=strip_residual_fragments,
                deduplicate_identical=deduplicate_identical,
            )
            with tempfile.TemporaryDirectory(prefix="steel_cl_") as cl_tmp:
                cfg, raw_plates, ver, unique_plates, instance_mapping, _, _, pre_post_ver = (
                    _extract_and_verify_cheson_closed_loop(
                        dstv=dstv,
                        base_cfg=base_cfg,
                        bom=bom,
                        bom_row=bom_row,
                        pdf_analysis=pdf_analysis,
                        noted_dims=noted_dims,
                        a_pdf=a_pdf,
                        cached_pdf_views=cached_pdf_views,
                        use_nesting_layers=use_nesting_layers,
                        tmp_dir=Path(cl_tmp),
                    )
                )
            verifications.append(asdict(ver))

            if piece_key in excel_qty_map and resolved_xlsx is not None:
                cheson_qty = excel_qty_map[piece_key]
                qty_source = resolved_xlsx.name
            elif bom_row is not None:
                cheson_qty = max(1, bom_row.qty_per_assembly) * max(1, assembly_qty)
                qty_source = "STYKKLISTE"
            else:
                cheson_qty = max(1, dstv.header.quantity)
                qty_source = "NC1"

            badge = "✅" if ver.passed else "⚠"
            _log(
                f"🧩 {nc1_path.name} ({dstv.header.profile_str}) -> {ver.detected_type_label} "
                f"[{len(raw_plates)} fizice -> {len(unique_plates)} unice, {cheson_qty} chesoane ({qty_source})]"
            )
            _log(
                f"   {badge} Gabarit: L={ver.l_assembled_mm:.3f} mm (țintă {ver.l_gabarit_target_mm:.3f} [{ver.l_source_label}], "
                f"NC1 brut {ver.l_nc1_raw_mm:.2f}) | H={ver.h_assembled_mm:.3f}/{ver.h_target_mm:.3f} | "
                f"B={ver.w_assembled_mm:.3f}/{ver.w_target_mm:.3f} | Eroare max={ver.max_gabarit_error_mm:.4f} mm"
            )

            unique_audits: list[DrawnDxfAudit] = []
            unique_rings: dict[str, Ring] = {}
            current_plate_rows: list[dict[str, Any]] = []

            for plate in unique_plates:
                files, rel, dxf_aud, rr = _export_plate_cad(
                    plate,
                    out_dir=out_root,
                    export_format=export_format,
                    use_nesting_layers=use_nesting_layers,
                    include_marking=include_marking,
                    organize_by_thickness=organize_by_thickness,
                    oda_path=effective_oda,
                    log_cb=_log,
                )
                unique_audits.append(dxf_aud)
                unique_rings[plate.name] = rr
                written_files.extend(files)
                nc1_plate_count += 1
                minx, miny, maxx, maxy = plate.bounds
                dim_l, dim_w = round(maxx - minx, 4), round(maxy - miny, 4)
                role_label = _ROLE_LABELS_RO.get(plate.role, plate.role)
                tot_qty = plate.qty_per_cheson * cheson_qty
                res_msg = (
                    f" | 🧹 {plate.removed_residuals} reziduuri ≤ {plate.thickness:g} mm eliminate"
                    if plate.removed_residuals > 0
                    else ""
                )
                holes_msg = (
                    f" | 🕳️ {dxf_aud.holes_matched}/{dxf_aud.holes_expected} găuri"
                    if dxf_aud.holes_expected > 0
                    else ""
                )
                merge_msg = (
                    f" | Identice={','.join(plate.merged_sources)}"
                    if len(plate.merged_sources) > 1
                    else ""
                )
                _log(
                    f"   📄 {rel} [{plate.role_code} {role_label}] ({dim_l:.3f} × {dim_w:.3f} mm, "
                    f"T{int(round(plate.thickness))}, ×{plate.qty_per_cheson}/cheson -> {tot_qty} buc, "
                    f"Layer={','.join(dxf_aud.entity_layers) or '0'}){res_msg}{holes_msg}{merge_msg}"
                )
                entry_obj = ManifestEntry(
                    pos_no=dstv.header.piece_id or nc1_path.stem,
                    plate_name=plate.name,
                    profile=dstv.header.profile_str,
                    material=plate.material,
                    thickness_mm=plate.thickness,
                    width_mm=dim_w,
                    length_mm=dim_l,
                    qty_per_assembly=plate.qty_per_cheson,
                    assembly_qty=cheson_qty,
                    total_qty=tot_qty,
                    classification="COMPOUND_NC1" if len(raw_plates) > 1 else "SINGLE_NC1",
                    status="VERIFIED_NC1",
                    source_file=nc1_path.name,
                    output_folder=plate.folder if organize_by_thickness else "",
                    output_file=rel,
                    notes=(
                        f"type={ver.detected_type}, role={role_label} ({plate.role_code}), "
                        f"identice={','.join(plate.merged_sources)}, "
                        f"layers={','.join(dxf_aud.entity_layers)}, dim_err={dxf_aud.dim_err_mm:.4f}mm"
                    ),
                    role_code=plate.role_code,
                    role_label=role_label,
                    qty_per_cheson=plate.qty_per_cheson,
                    qty_source=qty_source,
                )
                entries.append(entry_obj)
                current_plate_rows.append(asdict(entry_obj))

            dxf_audits, re_read_rings = _expand_audits_for_3d_sim(
                raw_plates, unique_plates, instance_mapping, unique_audits, unique_rings
            )

            post_ver = verify_post_draw_and_3d(
                dstv=dstv,
                cfg=cfg,
                plates=raw_plates,
                dxf_audits=dxf_audits,
                re_read_rings=re_read_rings,
                bom=bom,
                bom_row=bom_row,
                pdf_dims=pdf_dims,
                pdf_analysis=pdf_analysis,
                a_pdf_path=a_pdf,
                cached_pdf_views=cached_pdf_views,
            )
            post_ver.refinement_iterations = pre_post_ver.refinement_iterations
            post_ver.refinement_notes = list(pre_post_ver.refinement_notes)
            post_ver.snapped_noted_dims = list(pre_post_ver.snapped_noted_dims)
            post_draw_audits.append(asdict(post_ver))
            tcl = format_cheson_tcl_summary(current_plate_rows)
            cheson_summaries.append(
                {
                    "cheson_id": dstv.header.piece_id or nc1_path.stem,
                    "folder": search_root.name,
                    "profile": dstv.header.profile_str,
                    "detected_type_label": ver.detected_type_label,
                    "cheson_qty": cheson_qty,
                    "qty_source": qty_source,
                    "talpa_str": tcl["talpa_str"],
                    "capac_str": tcl["capac_str"],
                    "laterala_str": tcl["laterala_str"],
                    "full_summary": tcl["full_summary"],
                    "unique_plates_count": len(unique_plates),
                    "total_pieces_count": sum(int(r["total_qty"]) for r in current_plate_rows),
                    "overlay_pct": (post_ver.pdf_overlay.overall_overlay_pct if post_ver.pdf_overlay else None),
                    "status": "PASS" if (ver.passed and post_ver.passed) else "CHECK",
                }
            )
            _log(f"   📊 Cantități T/C/L ({cheson_qty} chesoane [{qty_source}]): {tcl['full_summary']}")
            if post_ver.snapped_noted_dims:
                _log(f"   📐 Cote PDF blocate (≤0.001 mm): {', '.join(post_ver.snapped_noted_dims)}")
            if post_ver.refinement_iterations > 1:
                _log(f"   🔄 Auto-corecție închisă: {post_ver.refinement_iterations} iterații ({post_ver.refinement_notes[-1]})")
            s3 = post_ver.sim_3d
            p_badge = "✅ PASS" if post_ver.passed else "⚠ CHECK"
            _log(
                f"   🔎 REVERIFICARE POST-DESENARE ({p_badge}): DXF curat={post_ver.all_dxf_clean} | "
                f"Layer 0 pur={post_ver.all_layer_zero} | Err cote={post_ver.nc1_max_dim_err_mm:.4f} mm | "
                f"Err găuri={post_ver.nc1_max_hole_err_mm:.4f} mm"
            )
            if post_ver.pdf_overlay is not None:
                ov = post_ver.pdf_overlay
                ov_badge = "✅" if ov.passed else "⚠"
                sc_s = f"1:{int(round(ov.top_scale or 25))}"
                _log(
                    f"   🎯 SUPRAPUNERE PDF LA SCARĂ ({sc_s}): {ov_badge} {ov.overall_overlay_pct:.2f}% "
                    f"(TOP={ov.top_overlay_pct:.2f}% [err {ov.top_mean_err_mm:.2f} mm], "
                    f"FRONT={ov.front_overlay_pct:.2f}% [err {ov.front_mean_err_mm:.2f} mm], "
                    f"Găuri PDF={ov.holes_pdf_matched}/{ov.holes_pdf_expected})"
                )
            if post_ver.pdf_drawing_name:
                pdf_preview_dims = ", ".join(f"{v:g}" for v in post_ver.pdf_dims_confirmed[:10])
                bom_err_s = f"{post_ver.pdf_bom_length_err_mm:.4f} mm" if post_ver.pdf_bom_length_err_mm is not None else "-"
                holes_call_s = (
                    f" | Găuri PDF=[{', '.join(f'{q}xØ{d:g}' for q, d in post_ver.pdf_hole_callouts[:5])}]"
                    if post_ver.pdf_hole_callouts
                    else ""
                )
                prof_call_s = (
                    f" | Profil PDF={post_ver.pdf_profile_callouts[0]}"
                    if post_ver.pdf_profile_callouts
                    else ""
                )
                _log(
                    f"   📋 PDF ({post_ver.pdf_drawing_name} [{post_ver.pdf_text_source}]): ΔL STYKKLISTE={bom_err_s}{prof_call_s}{holes_call_s} | "
                    f"Cote confirmate ({len(post_ver.pdf_dims_confirmed)}): [{pdf_preview_dims}]"
                )
            iou_str = ", ".join(f"{k}={v:.2f}%" for k, v in s3.face_iou_pct.items())
            mass_str = (
                f"{s3.mass_3d_kg:.2f} kg (BOM {s3.mass_bom_kg:.2f} kg, Δ{s3.mass_diff_pct:.2f}%)"
                if s3.mass_bom_kg
                else f"{s3.mass_3d_kg:.2f} kg"
            )
            _log(
                f"   🧊 SIMULARE ASAMBLARE 3D: {s3.l_3d_mm:.3f} × {s3.b_3d_mm:.3f} × {s3.h_3d_mm:.3f} mm | "
                f"Coliziuni={s3.max_interference_mm:.3f} mm | Rost sudură={s3.max_seam_gap_mm:.3f} mm | "
                f"Fidelitate 3D vs NC1={s3.overall_3d_similarity_pct:.2f}% ({iou_str}) | "
                f"Găuri 3D={s3.holes_3d_matched}/{s3.holes_3d_expected} | Masă={mass_str}"
            )
            if post_ver.issues:
                for iss in post_ver.issues[:5]:
                    _log(f"      ⚠ {iss}")
        except Exception as exc:
            err = f"Eroare la procesarea {nc1_path.name}: {exc}"
            errors.append(err)
            _log(f"   ❌ {err}")
        done_steps += 1
        _prog(done_steps / total_steps)

    # Optional W-drawing PDF conversion via PdfToCadConverter
    if scope_mode in ("cheson_and_w", "full_assembly") and w_map:
        if PdfToCadConverter is None or not PdfToCadConverter.is_available():
            msg = "PdfToCadConverter indisponibil pentru W-drawings."
            errors.append(msg)
            _log(f"⚠ {msg}")
        else:
            conv = PdfToCadConverter(oda_path=effective_oda)
            w_req_fmt = "dxf" if export_format in ("dxf", "both") else "dwg"
            for key, w_pdf in sorted(w_map.items()):
                if key in processed_keys:
                    done_steps += 1
                    continue
                processed_keys.add(key)
                bom_row = bom_by_key.get(key)
                t_mm: Optional[float] = None
                w_mm: Optional[float] = None
                l_mm: Optional[float] = None
                pos_no = bom_row.pos_no if bom_row else w_pdf.stem.split("-")[0].strip()
                material = bom_row.material if bom_row else ""
                profile = bom_row.profile if bom_row else ""
                qty_per = bom_row.qty_per_assembly if bom_row else 1

                if bom_row:
                    pl = parse_plate_profile(bom_row.profile)
                    if pl:
                        t_mm, w_mm = pl
                    l_mm = bom_row.length_mm

                if t_mm is None and pymupdf is not None:
                    try:
                        d_tmp = pymupdf.open(str(w_pdf))
                        meta = conv.extract_drawing_metadata(d_tmp[0], str(w_pdf))
                        d_tmp.close()
                        if meta.get("profil_thickness"):
                            t_mm = float(meta["profil_thickness"])
                        if meta.get("profil_width"):
                            w_mm = float(meta["profil_width"])
                        l_mm = l_mm or meta.get("lengde")
                    except Exception:
                        pass

                folder_name = f"T{int(round(t_mm))}" if t_mm else "W_DRAWING"
                target_dir = (out_root / folder_name) if organize_by_thickness else out_root
                target_dir.mkdir(parents=True, exist_ok=True)
                clean_nm = clean_piece_name(pos_no, lowercase=True)

                ok, out_cad, conv_msg = conv.convert_single_pdf(
                    str(w_pdf),
                    output_dir=str(target_dir),
                    output_format=w_req_fmt,
                    only_front_view=True,
                    calibrate_1to1=True,
                    custom_name=clean_nm,
                )
                if ok and out_cad:
                    out_p = Path(out_cad)
                    written_files.append(out_p)
                    if export_format == "both" and out_p.suffix.lower() == ".dxf" and effective_oda:
                        try:
                            dwg_p = out_p.with_suffix(".dwg")
                            convert_dxf_to_dwg(out_p, dwg_p, oda_path=effective_oda)
                            written_files.append(dwg_p)
                            out_p = dwg_p
                        except Exception as exc_oda:
                            _log(f"  ⚠ Eroare ODA pentru W-drawing {w_pdf.name}: {exc_oda}")
                    w_count += 1
                    rel = f"{folder_name}/{out_p.name}" if organize_by_thickness else out_p.name
                    _log(f"📐 W-drawing {w_pdf.name} -> {rel} ({conv_msg})")
                    entries.append(
                        ManifestEntry(
                            pos_no=pos_no,
                            plate_name=clean_nm,
                            profile=profile,
                            material=material,
                            thickness_mm=t_mm,
                            width_mm=w_mm,
                            length_mm=l_mm,
                            qty_per_assembly=qty_per,
                            assembly_qty=assembly_qty,
                            total_qty=qty_per * max(1, assembly_qty),
                            classification="HAS_W_DRAWING",
                            status="CONVERTED_W_DRAWING",
                            source_file=w_pdf.name,
                            output_folder=folder_name if organize_by_thickness else "",
                            output_file=rel,
                            notes=conv_msg,
                        )
                    )
                else:
                    err = f"W-drawing {w_pdf.name} eșuat: {conv_msg}"
                    errors.append(err)
                    _log(f"   ❌ {err}")
                done_steps += 1
                _prog(done_steps / total_steps)

    # Optional BOM rectangular plates from A-drawing.pdf
    if scope_mode == "full_assembly" and bom:
        for row in bom.rows:
            key = clean_piece_name(row.pos_no, lowercase=True)
            if key in processed_keys:
                done_steps += 1
                continue
            processed_keys.add(key)

            pl = parse_plate_profile(row.profile)
            if pl is not None and row.length_mm is not None and row.length_mm > 0.0:
                t_mm, w_mm = pl
                l_mm = float(row.length_mm)
                dim_x = max(l_mm, w_mm)
                dim_y = min(l_mm, w_mm)
                clean_nm = clean_piece_name(row.pos_no, lowercase=True)
                rect_plate = ExtractedPlate(
                    name=clean_nm,
                    folder=f"T{int(round(t_mm))}",
                    role="bom_plate",
                    face="bom",
                    thickness=t_mm,
                    piece_id=row.pos_no,
                    piece_clean=clean_nm,
                    material=row.material,
                    quantity=row.qty_per_assembly * max(1, assembly_qty),
                    exterior=[(0.0, 0.0), (dim_x, 0.0), (dim_x, dim_y), (0.0, dim_y)],
                    source_bounds=(0.0, 0.0, dim_x, dim_y),
                )
                files, rel, _, _ = _export_plate_cad(
                    rect_plate,
                    out_dir=out_root,
                    export_format=export_format,
                    use_nesting_layers=use_nesting_layers,
                    include_marking=include_marking,
                    organize_by_thickness=organize_by_thickness,
                    oda_path=effective_oda,
                    log_cb=_log,
                )
                written_files.extend(files)
                bom_plate_count += 1
                _log(f"📄 BOM placă {row.pos_no} ({row.profile} x {l_mm:.0f}) -> {rel} [DRAFT_RECT_CHECK]")
                entries.append(
                    ManifestEntry(
                        pos_no=row.pos_no,
                        plate_name=clean_nm,
                        profile=row.profile,
                        material=row.material,
                        thickness_mm=t_mm,
                        width_mm=dim_y,
                        length_mm=dim_x,
                        qty_per_assembly=row.qty_per_assembly,
                        assembly_qty=assembly_qty,
                        total_qty=row.qty_per_assembly * max(1, assembly_qty),
                        classification="BOM_PLATE_ONLY",
                        status="DRAFT_RECT_CHECK",
                        source_file=bom.source_pdf.name,
                        output_folder=rect_plate.folder if organize_by_thickness else "",
                        output_file=rel,
                        notes="Placă dreptunghiulară sintetizată din BOM (verificați A-drawing pentru teșituri/găuri)",
                    )
                )
            else:
                skipped_count += 1
                entries.append(
                    ManifestEntry(
                        pos_no=row.pos_no,
                        plate_name=clean_piece_name(row.pos_no, lowercase=True),
                        profile=row.profile,
                        material=row.material,
                        thickness_mm=None,
                        width_mm=None,
                        length_mm=row.length_mm,
                        qty_per_assembly=row.qty_per_assembly,
                        assembly_qty=assembly_qty,
                        total_qty=row.qty_per_assembly * max(1, assembly_qty),
                        classification="NON_PLATE_OTHER",
                        status="SUB_ASSEMBLY_SKIPPED",
                        source_file=bom.source_pdf.name,
                        output_folder="",
                        output_file="",
                        notes="Subansamblu sau profil fără .nc1/W-drawing",
                    )
                )
            done_steps += 1
            _prog(done_steps / total_steps)

    manifest_path: Optional[Path] = None
    csv_path: Optional[Path] = None
    excel_report_path: Optional[Path] = None
    if generate_reports:
        status_counts: dict[str, int] = {}
        folder_counts: dict[str, int] = {}
        for e in entries:
            status_counts[e.status] = status_counts.get(e.status, 0) + 1
            if e.output_folder:
                folder_counts[e.output_folder] = folder_counts.get(e.output_folder, 0) + 1

        manifest_path = out_root / "manifest.json"
        payload = {
            "assembly": {
                "assembly_id": assembly_id,
                "assembly_quantity": assembly_qty,
                "source_path": str(src),
                "scope_mode": scope_mode,
                "preset": preset,
                "exact_gabarit_lock": exact_gabarit_lock,
                "strip_residual_fragments": strip_residual_fragments,
                "deduplicate_identical": deduplicate_identical,
                "use_extra_layers": use_nesting_layers,
                "export_format": export_format,
            },
            "cheson_summaries": cheson_summaries,
            "verifications": verifications,
            "post_draw_audits": post_draw_audits,
            "summary": {
                "total_entries": len(entries),
                "by_status": status_counts,
                "by_thickness_folder": dict(sorted(folder_counts.items())),
            },
            "items": [asdict(e) for e in entries],
        }
        manifest_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

        csv_path = out_root / "report.csv"
        fieldnames = [
            "pos_no", "plate_name", "role_code", "role_label", "profile", "material",
            "thickness_mm", "width_mm", "length_mm", "qty_per_cheson",
            "qty_per_assembly", "assembly_qty", "total_qty", "qty_source",
            "classification", "status", "source_file", "output_folder", "output_file", "notes",
        ]
        with csv_path.open("w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames)
            writer.writeheader()
            for e in entries:
                writer.writerow(asdict(e))

        if _write_excel_report:
            excel_report_path = write_cheson_quantity_report_xlsx(
                out_root / "Raport_Piese_Cheson.xlsx",
                cheson_summaries=cheson_summaries,
                plate_entries=[asdict(e) for e in entries],
            )
            if excel_report_path is not None:
                _log(f"📊 Raport Excel generat: {excel_report_path}")

    _prog(1.0)
    _log(
        f"\n🏁 Finalizat: {nc1_plate_count} piese unice cheson (.nc1), "
        f"{w_count} W-drawings, {bom_plate_count} plăci BOM, {skipped_count} sărite."
    )
    return {
        "plates_generated": nc1_plate_count,
        "w_drawings_converted": w_count,
        "bom_plates_generated": bom_plate_count,
        "skipped": skipped_count,
        "cheson_summaries": cheson_summaries,
        "verifications": verifications,
        "post_draw_audits": post_draw_audits,
        "files": [str(f) for f in written_files],
        "manifest_path": str(manifest_path) if manifest_path else None,
        "csv_path": str(csv_path) if csv_path else None,
        "excel_report_path": str(excel_report_path) if excel_report_path else None,
        "dest_dir": str(out_root),
        "errors": errors,
        "_entries": [asdict(e) for e in entries],
    }


def run_cheson_pipeline(
    source_path: str | Path,
    dest_dir: str | Path,
    source_mode: str = "folder",
    scope_mode: str = "cheson_only",
    preset: str = "auto",
    export_format: str = "dwg_fallback",
    use_nesting_layers: bool = False,
    include_marking: bool = False,
    organize_by_thickness: bool = True,
    recursive: bool = False,
    generate_reports: bool = True,
    exact_gabarit_lock: bool = True,
    override_length_mm: Optional[float] = None,
    override_web_height_mm: Optional[float] = None,
    oda_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    strip_residual_fragments: bool = True,
    save_in_source_subfolders: bool = False,
    deduplicate_identical: bool = True,
    cheson_report_xlsx: Optional[str | Path] = None,
) -> dict[str, Any]:
    """Run box-beam (.nc1) extraction on a single file, single folder, or multi-folder cheson batch."""
    def _log(msg: str) -> None:
        if log_callback:
            log_callback(msg)

    def _prog(val: float) -> None:
        if progress_callback:
            progress_callback(max(0.0, min(1.0, val)))

    src = Path(source_path).resolve()
    out_root = Path(dest_dir).resolve()

    if source_mode != "single_file" and src.is_dir():
        batch_folders = discover_cheson_batch_folders(src, recursive=recursive)
        is_multi = len(batch_folders) > 1 or (len(batch_folders) == 1 and batch_folders[0] != src)
        if is_multi:
            report_root = src if save_in_source_subfolders else out_root
            report_root.mkdir(parents=True, exist_ok=True)
            resolved_report_xlsx = find_cheson_report_xlsx(src, cheson_report_xlsx)
            _log(f"🗂️ BATCH AUTOMAT CHESON: {len(batch_folders)} foldere detectate în {src}")
            if resolved_report_xlsx is not None:
                _log(f"📊 Raport cantități chesoane detectat: {resolved_report_xlsx.name}")
            _log(
                "   Mod salvare: Direct în folderul fiecărui cheson"
                if save_in_source_subfolders
                else f"   Mod salvare: Mirror organizat în {out_root}"
            )

            all_verifications: list[dict[str, Any]] = []
            all_audits: list[dict[str, Any]] = []
            all_files: list[str] = []
            all_errors: list[str] = []
            all_entries: list[dict[str, Any]] = []
            all_cheson_summaries: list[dict[str, Any]] = []
            folder_summaries: list[dict[str, Any]] = []

            total_nc1_plates = 0
            total_w = 0
            total_bom = 0
            total_skipped = 0
            batch_pass_count = 0
            batch_check_count = 0

            n_folders = len(batch_folders)
            for idx, sub_dir in enumerate(batch_folders, start=1):
                rel_folder = str(sub_dir.relative_to(src)).replace("\\", "/")
                sub_out = sub_dir if save_in_source_subfolders else (out_root / sub_dir.relative_to(src))
                _log(f"\n═══ [{idx}/{n_folders}] Folder cheson: {rel_folder} ═══")

                def _sub_prog(p: float, i: int = idx - 1, n: int = n_folders) -> None:
                    _prog((i + max(0.0, min(1.0, p))) / max(1, n))

                try:
                    sub_res = _run_cheson_single_folder(
                        source_path=sub_dir,
                        dest_dir=sub_out,
                        source_mode="folder",
                        scope_mode=scope_mode,
                        preset=preset,
                        export_format=export_format,
                        use_nesting_layers=use_nesting_layers,
                        include_marking=include_marking,
                        organize_by_thickness=organize_by_thickness,
                        recursive=False,
                        generate_reports=generate_reports,
                        exact_gabarit_lock=exact_gabarit_lock,
                        override_length_mm=override_length_mm,
                        override_web_height_mm=override_web_height_mm,
                        oda_path=oda_path,
                        progress_callback=_sub_prog,
                        log_callback=_log,
                        strip_residual_fragments=strip_residual_fragments,
                        deduplicate_identical=deduplicate_identical,
                        cheson_report_xlsx=resolved_report_xlsx or cheson_report_xlsx,
                        _write_excel_report=False,
                    )
                    sub_entries = sub_res.pop("_entries", [])
                    sub_vers = sub_res.get("verifications") or []
                    sub_audits = sub_res.get("post_draw_audits") or []
                    sub_errs = sub_res.get("errors") or []
                    sub_ch_sums = sub_res.get("cheson_summaries") or []

                    for cs in sub_ch_sums:
                        cs["folder"] = rel_folder
                        all_cheson_summaries.append(cs)
                    for v in sub_vers:
                        v["folder"] = rel_folder
                    for a in sub_audits:
                        a["folder"] = rel_folder
                    for e in sub_entries:
                        e_with_folder = {"cheson_folder": rel_folder, **e}
                        if e.get("output_file") and not save_in_source_subfolders:
                            e_with_folder["output_file"] = f"{rel_folder}/{e['output_file']}"
                        all_entries.append(e_with_folder)

                    all_verifications.extend(sub_vers)
                    all_audits.extend(sub_audits)
                    all_files.extend(sub_res.get("files") or [])
                    all_errors.extend(f"[{rel_folder}] {er}" for er in sub_errs)

                    total_nc1_plates += int(sub_res.get("plates_generated", 0))
                    total_w += int(sub_res.get("w_drawings_converted", 0))
                    total_bom += int(sub_res.get("bom_plates_generated", 0))
                    total_skipped += int(sub_res.get("skipped", 0))

                    folder_ok = (
                        not sub_errs
                        and all(v.get("passed", True) for v in sub_vers)
                        and all(a.get("passed", True) for a in sub_audits)
                    )
                    if folder_ok:
                        batch_pass_count += 1
                    else:
                        batch_check_count += 1

                    folder_summaries.append({
                        "folder": rel_folder,
                        "status": "PASS" if folder_ok else "CHECK",
                        "plates_generated": int(sub_res.get("plates_generated", 0)),
                        "errors_count": len(sub_errs),
                    })
                except Exception as exc:
                    batch_check_count += 1
                    err_msg = f"[{rel_folder}] Eroare folder: {exc}"
                    all_errors.append(err_msg)
                    _log(f"❌ {err_msg}")
                    folder_summaries.append({
                        "folder": rel_folder,
                        "status": "ERROR",
                        "plates_generated": 0,
                        "errors_count": 1,
                    })
                _prog(idx / max(1, n_folders))

            manifest_path: Optional[Path] = None
            csv_path: Optional[Path] = None
            excel_report_path: Optional[Path] = None
            if generate_reports:
                manifest_path = report_root / "batch_manifest.json"
                batch_payload = {
                    "batch_root": str(src),
                    "save_in_source_subfolders": save_in_source_subfolders,
                    "deduplicate_identical": deduplicate_identical,
                    "cheson_report_xlsx": resolved_report_xlsx.name if resolved_report_xlsx else None,
                    "folders_processed": n_folders,
                    "batch_pass_count": batch_pass_count,
                    "batch_check_count": batch_check_count,
                    "totals": {
                        "plates_generated": total_nc1_plates,
                        "w_drawings_converted": total_w,
                        "bom_plates_generated": total_bom,
                        "skipped": total_skipped,
                    },
                    "cheson_summaries": all_cheson_summaries,
                    "folders": folder_summaries,
                    "verifications": all_verifications,
                    "post_draw_audits": all_audits,
                    "items": all_entries,
                }
                manifest_path.write_text(json.dumps(batch_payload, indent=2, ensure_ascii=False), encoding="utf-8")

                csv_path = report_root / "batch_report.csv"
                fieldnames = [
                    "cheson_folder", "pos_no", "plate_name", "role_code", "role_label",
                    "profile", "material", "thickness_mm", "width_mm", "length_mm",
                    "qty_per_cheson", "qty_per_assembly", "assembly_qty", "total_qty",
                    "qty_source", "classification", "status", "source_file",
                    "output_folder", "output_file", "notes",
                ]
                with csv_path.open("w", newline="", encoding="utf-8-sig") as fh:
                    writer = csv.DictWriter(fh, fieldnames=fieldnames)
                    writer.writeheader()
                    for row_d in all_entries:
                        writer.writerow(row_d)

                excel_report_path = write_cheson_quantity_report_xlsx(
                    report_root / "Raport_Piese_Cheson.xlsx",
                    cheson_summaries=all_cheson_summaries,
                    plate_entries=all_entries,
                )

            _prog(1.0)
            if all_cheson_summaries:
                _log("\n📊 SUMAR GENERAL PIESE (T / C / L):")
                for cs in all_cheson_summaries:
                    _log(f"   • {cs['cheson_id']} (×{cs['cheson_qty']} [{cs['qty_source']}]): {cs['full_summary']}")
            if excel_report_path is not None:
                _log(f"📊 Raport Excel central generat: {excel_report_path}")
            _log(
                f"\n🏁 BATCH FINALIZAT: {n_folders} foldere cheson ({batch_pass_count} PASS, {batch_check_count} CHECK) | "
                f"{total_nc1_plates} piese unice NC1, {total_w} W-drawings, {total_bom} plăci BOM."
            )
            return {
                "is_multi_folder_batch": True,
                "folders_processed": n_folders,
                "batch_pass_count": batch_pass_count,
                "batch_check_count": batch_check_count,
                "plates_generated": total_nc1_plates,
                "w_drawings_converted": total_w,
                "bom_plates_generated": total_bom,
                "skipped": total_skipped,
                "cheson_summaries": all_cheson_summaries,
                "verifications": all_verifications,
                "post_draw_audits": all_audits,
                "files": all_files,
                "manifest_path": str(manifest_path) if manifest_path else None,
                "csv_path": str(csv_path) if csv_path else None,
                "excel_report_path": str(excel_report_path) if excel_report_path else None,
                "dest_dir": str(report_root),
                "errors": all_errors,
            }

    res = _run_cheson_single_folder(
        source_path=source_path,
        dest_dir=dest_dir,
        source_mode=source_mode,
        scope_mode=scope_mode,
        preset=preset,
        export_format=export_format,
        use_nesting_layers=use_nesting_layers,
        include_marking=include_marking,
        organize_by_thickness=organize_by_thickness,
        recursive=recursive,
        generate_reports=generate_reports,
        exact_gabarit_lock=exact_gabarit_lock,
        override_length_mm=override_length_mm,
        override_web_height_mm=override_web_height_mm,
        oda_path=oda_path,
        progress_callback=progress_callback,
        log_callback=log_callback,
        strip_residual_fragments=strip_residual_fragments,
        deduplicate_identical=deduplicate_identical,
        cheson_report_xlsx=cheson_report_xlsx,
        _write_excel_report=True,
    )
    res.pop("_entries", None)
    auds = res.get("post_draw_audits") or []
    vers = res.get("verifications") or []
    folder_ok = bool(
        not res.get("errors")
        and all(v.get("passed", True) for v in vers)
        and all(a.get("passed", True) for a in auds)
    )
    res["is_multi_folder_batch"] = False
    res["folders_processed"] = 1
    res["batch_pass_count"] = 1 if folder_ok else 0
    res["batch_check_count"] = 0 if folder_ok else 1
    return res
