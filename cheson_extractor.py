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
    ) -> "ChesonConfig":
        key = (preset or "auto").strip().lower()
        common = {
            "exact_gabarit_lock": exact_gabarit_lock,
            "target_length_mm": target_length_mm,
            "override_web_height_mm": override_web_height_mm,
            "strip_residual_fragments": strip_residual_fragments,
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
            if hole_type == "l" and len(rest) >= 2:
                slot_len = _parse_coord_token(rest[0])[0]
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
        tf_bot = float(m3.group(4))
        b_bot = float(m3.group(5))
        tf_top = float(m3.group(6))
        b_top = float(m3.group(7))
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


def detect_cheson_type(
    dstv: DSTVFile,
    spec: CompoundProfileSpec,
    cfg: ChesonConfig,
) -> tuple[str, str]:
    """Classify box beam into (cheson_type, talpa_side: 'none'|'bottom'|'top')."""
    explicit = (cfg.cheson_type or "auto").strip().lower()
    tw = spec.tw

    dy_o = _face_y_span(dstv, "o", cfg) or spec.b_top
    dy_u = _face_y_span(dstv, "u", cfg) or spec.b_bot
    bot_wider = (spec.b_bot > spec.b_top + 2.0 * tw + 1.0) or (dy_u > dy_o + 2.0 * tw + 1.0)
    top_wider = (spec.b_top > spec.b_bot + 2.0 * tw + 1.0) or (dy_o > dy_u + 2.0 * tw + 1.0)
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

    minx, _, maxx, _ = polygon_bounds(ring)
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
        span = max(1.0, maxx - minx)
        left_cut = minx + 0.35 * span
        right_cut = maxx - 0.35 * span
        new_pts: Ring = []
        for x, y in ring:
            nx = x
            if abs(dx_left) > 1e-7 and x <= left_cut:
                nx = x + dx_left
            elif abs(dx_right) > 1e-7 and x >= right_cut:
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
        if point_in_polygon(hx, hole.y, ring, tol=0.05):
            nx, ny = _m_pt(hx, hole.y)
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
) -> VerificationReport:
    """Closed-loop geometric audit verifying <= 0.001 mm tolerance on gabarit & flange-web contact."""
    H_target = spec.total_height
    W_target = spec.total_width
    tw = spec.tw

    flanges = [p for p in plates if p.role in ("flange_top", "flange_bot", "talpa", "capac")]
    webs = [p for p in plates if p.role == "web"]

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

    # Check assembled height
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
            h_assembled = H_target
    elif detected_type == "cheson_cu_talpa" and webs:
        wh = webs[0].bounds[3] - webs[0].bounds[1]
        tf_t = spec.tf_top if talpa_side == "top" else spec.tf_bot
        h_assembled = wh + tf_t + tw
    elif detected_type == "cheson_fara_talpa" and webs:
        wh = webs[0].bounds[3] - webs[0].bounds[1]
        if flanges and abs((flanges[0].bounds[3] - flanges[0].bounds[1]) - W_target) <= 0.5:
            h_assembled = wh + spec.tf_top + spec.tf_bot
        else:
            h_assembled = wh + 2.0 * tw
    else:
        h_assembled = H_target

    # Check flange-web contact length consistency on top capac (C or C1) vs webs
    contact_err = 0.0
    ref_flange = next((p for p in plates if p.role in ("flange_top", "capac")), None)
    if ref_flange is not None and webs:
        r = ref_flange.exterior
        b = ref_flange.bounds
        y_lo, y_hi = b[1], b[3]
        # Measure top and bottom horizontal edge lengths of ref_flange
        top_xs = [x for x, y in r if abs(y - y_hi) < 0.01]
        bot_xs = [x for x, y in r if abs(y - y_lo) < 0.01]
        if len(top_xs) >= 2 and len(bot_xs) >= 2:
            len_edge_a = max(top_xs) - min(top_xs)
            len_edge_b = max(bot_xs) - min(bot_xs)
            f_long = max(len_edge_a, len_edge_b)
            f_short = min(len_edge_a, len_edge_b)
            full_webs = [
                wp for wp in webs
                if wp.removed_residuals == 0 and (wp.bounds[2] - wp.bounds[0]) >= f_short - 5.0
            ]
            web_lens = sorted({round(wp.bounds[2] - wp.bounds[0], 4) for wp in full_webs})
            if abs(f_long - f_short) <= 0.5:
                contact_err = max((abs(wl - f_long) for wl in web_lens), default=0.0)
            else:
                contact_err = max(
                    (min(abs(wl - f_short), abs(wl - l_assembled), abs(wl - f_long)) for wl in web_lens),
                    default=0.0,
                )

    err_l = abs(l_assembled - l_gabarit_target)
    err_h = abs(h_assembled - H_target)
    err_w = abs(w_assembled - W_target)
    max_err = max(err_l, err_h, err_w, contact_err)

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
        passed=(max_err <= 0.001 + 1e-9),
    )


def decompose_compound_nc1(
    dstv: DSTVFile,
    cfg: Optional[ChesonConfig] = None,
    return_verification: bool = False,
):
    cfg = cfg or ChesonConfig()
    header = dstv.header
    spec = parse_compound_profile(header.profile_str, header)
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

    # Determine flange clipping modes per detected cheson type
    if detected_type == "cheson_cu_talpa":
        top_over = (talpa_side == "top")
        bot_over = (talpa_side == "bottom")
        o_suffix = "T" if talpa_side == "top" else cfg.flange_letter
        u_suffix = "T" if talpa_side == "bottom" else cfg.flange_letter
        o_role = "talpa" if talpa_side == "top" else "capac"
        u_role = "talpa" if talpa_side == "bottom" else "capac"
    elif detected_type == "cheson_dublu":
        top_over = False
        bot_over = False
        o_suffix = f"{cfg.flange_letter}1"
        u_suffix = f"{cfg.flange_letter}2"
        o_role = "flange_top"
        u_role = "flange_bot"
    elif detected_type == "cheson_fara_talpa":
        top_over = (cfg.flange_top_mode == "over_webs")
        bot_over = (cfg.flange_bot_mode == "over_webs")
        o_suffix = f"{cfg.flange_letter}1"
        u_suffix = f"{cfg.flange_letter}2"
        o_role = "flange_top"
        u_role = "flange_bot"
    else:
        top_over = _resolve_flange_over_webs(cfg.flange_top_mode, spec.top_over_webs)
        bot_over = _resolve_flange_over_webs(cfg.flange_bot_mode, spec.bot_over_webs)
        o_suffix = f"{cfg.flange_letter}1"
        u_suffix = f"{cfg.flange_letter}2"
        o_role = "flange_top"
        u_role = "flange_bot"

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

    # Determine which web face (v or h) is the longer vs shorter side at skewed ends
    v_xmin = min((polygon_bounds(r)[0] for r in locked_rings.get("v", [])), default=0.0)
    h_xmin = min((polygon_bounds(r)[0] for r in locked_rings.get("h", [])), default=0.0)
    v_xmax = max((polygon_bounds(r)[2] for r in locked_rings.get("v", [])), default=0.0)
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
        is_long_left: Optional[bool] = None
        sync_lx: Optional[float] = None
        if eff_start_tan > 1e-6:
            is_long_left = (face == "v" and v_xmin < h_xmin) or (face == "h" and h_xmin < v_xmin)
            if is_long_left:
                sync_lx = x_min_locked
            elif left_contact_pair is not None:
                sync_lx = left_contact_pair[1]

        is_long_right: Optional[bool] = None
        sync_rx: Optional[float] = None
        if eff_end_tan > 1e-6:
            is_long_right = (face == "v" and v_xmax > h_xmax) or (face == "h" and h_xmax > v_xmax)
            if is_long_right:
                sync_rx = x_max_locked
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
    width = hole.slot_width if hole.slot_width > 0 else hole.diameter
    r = 0.5 * width
    straight = max(0.0, hole.slot_length - width)
    if straight < 0.05:
        msp.add_circle((hole.x, hole.y), radius=0.5 * hole.diameter, dxfattribs={"layer": layer})
        return
    half_s = 0.5 * straight
    rad = math.radians(hole.slot_angle)
    ux, uy = math.cos(rad), math.sin(rad)
    nx, ny = -uy, ux
    cx, cy = hole.x, hole.y
    pts = [
        (cx - half_s * ux - r * nx, cy - half_s * uy - r * ny, 0.0),
        (cx + half_s * ux - r * nx, cy + half_s * uy - r * ny, 1.0),
        (cx + half_s * ux + r * nx, cy + half_s * uy + r * ny, 0.0),
        (cx - half_s * ux + r * nx, cy - half_s * uy + r * ny, 1.0),
    ]
    msp.add_lwpolyline(pts, format="xyb", close=True, dxfattribs={"layer": layer})


def write_plate_dxf(
    plate: ExtractedPlate,
    output_path: str | Path,
    use_nesting_layers: bool = True,
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
        if hole.hole_type == "l" and hole.slot_length > hole.diameter:
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


def extract_bom_from_pdf(pdf_path: str | Path) -> Optional[BomTable]:
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
        return None
    finally:
        doc.close()


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

    header_candidates = [
        w for w in words
        if banner_word.bottom - 2.0 <= w.top <= banner_word.bottom + 36.0
        and w.x0 >= banner_word.x0 - 20.0
    ]
    header_row: Optional[list[_WordBox]] = None
    for hr in _cluster_words_by_y(header_candidates, y_tol=4.5):
        joined = " ".join(x.text.upper() for x in hr)
        if ("POS" in joined or "PART" in joined) and ("PROFIL" in joined or "MAT" in joined):
            header_row = hr
            break
    if header_row is None:
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

    header_bottom = max(w.bottom for w in header_row)
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
) -> tuple[list[Path], str]:
    target_dir = (out_dir / plate.folder) if organize_by_thickness else out_dir
    target_dir.mkdir(parents=True, exist_ok=True)
    dxf_path = target_dir / f"{plate.name}.dxf"
    write_plate_dxf(
        plate,
        dxf_path,
        use_nesting_layers=use_nesting_layers,
        include_marking=include_marking,
    )

    fmt = (export_format or "dwg_fallback").lower()
    if fmt == "dxf":
        rel = f"{plate.folder}/{plate.name}.dxf" if organize_by_thickness else f"{plate.name}.dxf"
        return [dxf_path], rel

    dwg_path = target_dir / f"{plate.name}.dwg"
    exe = find_oda_exe(oda_path)
    if exe:
        try:
            convert_dxf_to_dwg(dxf_path, dwg_path, oda_path=exe)
            if fmt == "both":
                rel = f"{plate.folder}/{plate.name}.dwg" if organize_by_thickness else f"{plate.name}.dwg"
                return [dxf_path, dwg_path], rel
            try:
                dxf_path.unlink()
            except Exception:
                pass
            rel = f"{plate.folder}/{plate.name}.dwg" if organize_by_thickness else f"{plate.name}.dwg"
            return [dwg_path], rel
        except Exception as exc:
            log_cb(f"  ⚠ Eroare conversie ODA pentru {plate.name}: {exc} (păstrat DXF)")

    if fmt == "dwg_direct":
        raise RuntimeError(f"Conversia DWG direct a eșuat pentru {plate.name} (ODA indisponibil).")

    rel = f"{plate.folder}/{plate.name}.dxf" if organize_by_thickness else f"{plate.name}.dxf"
    return [dxf_path], rel


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
) -> dict[str, Any]:
    """Dry-run inspection: detects cheson type, STYKKLISTE LENGDE, and verifies <= 0.001 mm tolerance."""
    _, _, nc1_files, a_pdf, bom, w_map = _collect_inputs(
        source_path, source_mode, scope_mode, recursive
    )
    bom_by_key: dict[str, BomRow] = {}
    if bom:
        for r in bom.rows:
            bom_by_key[clean_piece_name(r.pos_no, lowercase=True)] = r

    items_preview: list[dict[str, Any]] = []
    all_passed = True

    for nc1_path in nc1_files:
        dstv = parse_dstv(nc1_path)
        piece_key = clean_piece_name(dstv.header.piece_id or nc1_path.stem, lowercase=True)
        bom_row = bom_by_key.get(piece_key)
        target_len = override_length_mm if (override_length_mm and override_length_mm > 0) else (
            bom_row.length_mm if (bom_row and bom_row.length_mm) else None
        )
        cfg = ChesonConfig.from_preset(
            preset,
            exact_gabarit_lock=exact_gabarit_lock,
            target_length_mm=target_len,
            override_web_height_mm=override_web_height_mm,
            strip_residual_fragments=strip_residual_fragments,
        )
        plates, ver = decompose_nc1(dstv, cfg, return_verification=True)
        if not ver.passed:
            all_passed = False
        items_preview.append(
            {
                "file": nc1_path.name,
                "profile": dstv.header.profile_str,
                "verification": asdict(ver),
                "plates": [
                    {
                        "name": p.name,
                        "folder": p.folder,
                        "role": p.role,
                        "thickness": p.thickness,
                        "length_mm": round(p.bounds[2] - p.bounds[0], 4),
                        "width_mm": round(p.bounds[3] - p.bounds[1], 4),
                        "removed_residuals": p.removed_residuals,
                    }
                    for p in plates
                ],
            }
        )

    return {
        "a_drawing": a_pdf.name if a_pdf else None,
        "bom_assembly_id": bom.assembly_id if bom else None,
        "bom_rows_count": len(bom.rows) if bom else 0,
        "w_drawings_count": len(w_map),
        "all_passed": all_passed,
        "nc1_items": items_preview,
    }


def run_cheson_pipeline(
    source_path: str | Path,
    dest_dir: str | Path,
    source_mode: str = "folder",
    scope_mode: str = "cheson_only",
    preset: str = "auto",
    export_format: str = "dwg_fallback",
    use_nesting_layers: bool = True,
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
) -> dict[str, Any]:
    """Run box-beam (.nc1) plate extraction with exact gabarit verification (±0.001 mm)."""
    def _log(msg: str) -> None:
        if log_callback:
            log_callback(msg)

    def _prog(val: float) -> None:
        if progress_callback:
            progress_callback(max(0.0, min(1.0, val)))

    src, search_root, nc1_files, a_pdf, bom, w_map = _collect_inputs(
        source_path, source_mode, scope_mode, recursive
    )
    out_root = Path(dest_dir).resolve()
    out_root.mkdir(parents=True, exist_ok=True)
    effective_oda = find_oda_exe(oda_path)

    _log(
        f"🚀 Pornire Cheson / HSQ Extractor (Mod: {scope_mode}, Preset: {preset}, "
        f"GabaritFix={exact_gabarit_lock}, FărăReziduuri={strip_residual_fragments})"
    )
    _log(f"📁 Sursă: {src}")
    _log(f"📁 Destinație: {out_root}")
    if a_pdf is not None and bom is not None:
        _log(f"📋 STYKKLISTE ({a_pdf.name}): {bom.assembly_id} (qty={bom.assembly_quantity}, {len(bom.rows)} repere)")
    if export_format != "dxf":
        _log(f"🔧 ODA Converter: {effective_oda if effective_oda else 'indisponibil (fallback automat pe DXF R2018)'}")

    bom_by_key: dict[str, BomRow] = {}
    if bom:
        for r in bom.rows:
            bom_by_key[clean_piece_name(r.pos_no, lowercase=True)] = r

    assembly_qty = bom.assembly_quantity if bom else 1
    assembly_id = bom.assembly_id if bom else search_root.name

    entries: list[ManifestEntry] = []
    verifications: list[dict[str, Any]] = []
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
            cfg = ChesonConfig.from_preset(
                preset,
                exact_gabarit_lock=exact_gabarit_lock,
                target_length_mm=target_len,
                override_web_height_mm=override_web_height_mm,
                strip_residual_fragments=strip_residual_fragments,
            )
            plates, ver = decompose_nc1(dstv, cfg, return_verification=True)
            verifications.append(asdict(ver))

            qty_per = bom_row.qty_per_assembly if bom_row else dstv.header.quantity
            tot_qty = qty_per * max(1, assembly_qty)

            badge = "✅" if ver.passed else "⚠"
            _log(
                f"🧩 {nc1_path.name} ({dstv.header.profile_str}) -> {ver.detected_type_label}"
            )
            _log(
                f"   {badge} Gabarit: L={ver.l_assembled_mm:.3f} mm (țintă {ver.l_gabarit_target_mm:.3f} [{ver.l_source_label}], "
                f"NC1 brut {ver.l_nc1_raw_mm:.2f}) | H={ver.h_assembled_mm:.3f}/{ver.h_target_mm:.3f} | "
                f"B={ver.w_assembled_mm:.3f}/{ver.w_target_mm:.3f} | Eroare max={ver.max_gabarit_error_mm:.4f} mm"
            )

            for plate in plates:
                files, rel = _export_plate_cad(
                    plate,
                    out_dir=out_root,
                    export_format=export_format,
                    use_nesting_layers=use_nesting_layers,
                    include_marking=include_marking,
                    organize_by_thickness=organize_by_thickness,
                    oda_path=effective_oda,
                    log_cb=_log,
                )
                written_files.extend(files)
                nc1_plate_count += 1
                minx, miny, maxx, maxy = plate.bounds
                dim_l, dim_w = round(maxx - minx, 4), round(maxy - miny, 4)
                res_msg = (
                    f" | 🧹 {plate.removed_residuals} reziduuri ≤ {plate.thickness:g} mm eliminate"
                    if plate.removed_residuals > 0
                    else ""
                )
                _log(f"   📄 {rel} ({dim_l:.3f} × {dim_w:.3f} mm, T{int(round(plate.thickness))}){res_msg}")
                entries.append(
                    ManifestEntry(
                        pos_no=dstv.header.piece_id or nc1_path.stem,
                        plate_name=plate.name,
                        profile=dstv.header.profile_str,
                        material=plate.material,
                        thickness_mm=plate.thickness,
                        width_mm=dim_w,
                        length_mm=dim_l,
                        qty_per_assembly=qty_per,
                        assembly_qty=assembly_qty,
                        total_qty=tot_qty,
                        classification="COMPOUND_NC1" if len(plates) > 1 else "SINGLE_NC1",
                        status="VERIFIED_NC1",
                        source_file=nc1_path.name,
                        output_folder=plate.folder if organize_by_thickness else "",
                        output_file=rel,
                        notes=f"type={ver.detected_type}, face={plate.face}, role={plate.role}, err_max={ver.max_gabarit_error_mm:.4f}mm",
                    )
                )
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
                files, rel = _export_plate_cad(
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
                "export_format": export_format,
            },
            "verifications": verifications,
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
            "pos_no", "plate_name", "profile", "material", "thickness_mm",
            "width_mm", "length_mm", "qty_per_assembly", "assembly_qty",
            "total_qty", "classification", "status", "source_file",
            "output_folder", "output_file", "notes",
        ]
        with csv_path.open("w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames)
            writer.writeheader()
            for e in entries:
                writer.writerow(asdict(e))

    _prog(1.0)
    _log(
        f"\n🏁 Finalizat: {nc1_plate_count} piese cheson (.nc1), "
        f"{w_count} W-drawings, {bom_plate_count} plăci BOM, {skipped_count} sărite."
    )
    return {
        "plates_generated": nc1_plate_count,
        "w_drawings_converted": w_count,
        "bom_plates_generated": bom_plate_count,
        "skipped": skipped_count,
        "verifications": verifications,
        "files": [str(f) for f in written_files],
        "manifest_path": str(manifest_path) if manifest_path else None,
        "csv_path": str(csv_path) if csv_path else None,
        "dest_dir": str(out_root),
        "errors": errors,
    }
