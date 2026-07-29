"""
Freight Loader — 53' Dry Van Load Planning Tool
Imports Excel packing lists, calculates balanced load plans, exports floor plan PDF and crate labels.
"""

import sys
import math
import datetime
import json
import os
import re
from dataclasses import dataclass
from typing import Optional

SAVE_DIR = os.path.dirname(os.path.abspath(__file__))

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLabel, QSpinBox,
    QLineEdit, QGroupBox, QSplitter, QScrollArea, QMessageBox,
    QHeaderView, QCheckBox, QSizePolicy, QDialog, QDialogButtonBox,
    QFileDialog, QFormLayout, QAbstractItemView, QFrame, QComboBox
)
from PyQt6.QtCore import Qt, QRectF, QPointF, pyqtSignal
from PyQt6.QtGui import (
    QPainter, QColor, QPen, QBrush, QFont, QPageLayout, QPageSize, QPdfWriter
)

# ── Colors ───────────────────────────────────────────────────────────────────
TEAL_DARK  = QColor("#2a7d6e")
TEAL_MED   = QColor("#3dab96")
TEAL_LIGHT = QColor("#cdeae5")
TAN        = QColor("#c0a070")
GRAY_LINE  = QColor("#b0b0b0")
VOID_BG    = QColor("#f8f8f8")

# ── 53' Dry Van specs ────────────────────────────────────────────────────────
TRAILER = {
    "name": "53' Dry Van",
    "interior_length_in": 636,
    "interior_width_in":  98,
    "interior_height_in": 110,
    "tandem_from_nose_in": 492,   # trailer tandem center (~41 ft from kingpin)
    "steer_tare":  11000,
    "drive_tare":   9000,
    "steer_limit": 12000,
    "drive_limit": 34000,
    "trailer_limit": 34000,
    "gross_limit":  80000,
}


# ── Data model ────────────────────────────────────────────────────────────────
@dataclass
class FreightPiece:
    piece_id: int
    crate_label: str     # "Crate 9"
    part_number: str
    length_in: float
    width_in: float
    height_in: float
    weight_lbs: float
    pcs: int = 0
    notes: str = ""
    placement: str = "auto"

@dataclass
class ShipmentInfo:
    customer: str = ""
    shipment_name: str = ""
    trailer_num: str = ""
    bol_num: str = ""
    date: str = ""
    notes: str = ""

@dataclass
class PalletSlot:
    row: int
    col: int                     # 0 = Position A / Left, 1 = Position B / Right
    piece: Optional[FreightPiece]
    row_start_in: float          # inches from nose to front of this row
    is_center: bool = False

@dataclass
class LoadPlan:
    slots: list
    axle_weights: dict
    violations: list
    total_length_used_in: float
    override: bool = False
    excluded_ids: list = None


# ── Algorithm ─────────────────────────────────────────────────────────────────
def plan_load(pieces: list, override: bool = False, excluded_ids: list = None) -> LoadPlan:
    """
    Pair pieces heaviest+lightest for balanced rows, then order rows center-heavy
    so the heaviest rows sit near mid-trailer (best for drive/trailer axle balance).
    """
    if not pieces:
        return LoadPlan([], {}, [], 0.0, override, excluded_ids or [])

    # Separate center and auto pieces
    center_pieces = [p for p in pieces if p.placement == "center"]
    auto_pieces   = [p for p in pieces if p.placement != "center"]

    # Center pieces form their own row unit: (piece, None, True)
    center_units = [(p, None, True) for p in center_pieces]

    # Auto pieces are paired heaviest+lightest: (a, b, False)
    n = len(auto_pieces)
    sorted_p = sorted(auto_pieces, key=lambda p: p.weight_lbs, reverse=True)
    auto_units = []
    lo, hi = 0, n - 1
    while lo <= hi:
        a = sorted_p[lo]
        b = sorted_p[hi] if lo != hi else None
        auto_units.append((a, b, False))
        lo += 1
        hi -= 1

    # Combine all units, sort by total weight descending (heaviest unit first)
    all_units = center_units + auto_units
    all_units.sort(
        key=lambda u: u[0].weight_lbs + (u[1].weight_lbs if u[1] else 0),
        reverse=True
    )

    # Center-heavy row placement: heaviest unit → middle row, then alternate toward nose/tail
    n_rows = len(all_units)
    mid = n_rows // 2
    order = [mid]
    lo_i, hi_i = mid - 1, mid + 1
    while lo_i >= 0 or hi_i < n_rows:
        if hi_i < n_rows:
            order.append(hi_i)
            hi_i += 1
        if lo_i >= 0:
            order.append(lo_i)
            lo_i -= 1

    row_assignments = [None] * n_rows
    for dest, unit in zip(order, all_units):
        row_assignments[dest] = unit

    # Build slots (nose to tail) with cumulative row depths
    slots = []
    cursor = 0.0
    for row_idx, unit in enumerate(row_assignments):
        a_piece, b_piece, is_center = unit
        row_depth = max(a_piece.length_in, b_piece.length_in if b_piece else 0)
        if is_center:
            # Both col=0 and col=1 point to the same piece, both marked is_center=True
            slots.append(PalletSlot(row=row_idx, col=0, piece=a_piece,
                                    row_start_in=cursor, is_center=True))
            slots.append(PalletSlot(row=row_idx, col=1, piece=a_piece,
                                    row_start_in=cursor, is_center=True))
        else:
            # Alternate which side gets the heavier piece so lateral weight stays balanced
            if row_idx % 2 == 0:
                left_piece, right_piece = a_piece, b_piece
            else:
                left_piece, right_piece = b_piece, a_piece
            slots.append(PalletSlot(row=row_idx, col=0, piece=left_piece,
                                    row_start_in=cursor, is_center=False))
            slots.append(PalletSlot(row=row_idx, col=1, piece=right_piece,
                                    row_start_in=cursor, is_center=False))
        cursor += row_depth

    axle_weights, violations = _calc_axle_weights(slots, override)
    return LoadPlan(slots, axle_weights, violations, cursor, override, excluded_ids or [])


def _calc_axle_weights(slots: list, override: bool) -> tuple:
    """
    Beam model: trailer kingpin (nose=0) and trailer tandem as two supports.
    Trailer tandem reaction = Σ(weight × center_from_nose) / tandem_position.
    Pin weight goes to tractor 5th wheel → split between drive and steer axles.
    """
    tandem_pos = TRAILER["tandem_from_nose_in"]
    total_cargo = 0.0
    moment = 0.0

    for slot in slots:
        if slot.is_center and slot.col == 1:
            continue  # center piece already counted via col=0
        if slot.piece is None:
            continue
        center = slot.row_start_in + slot.piece.length_in / 2.0
        total_cargo += slot.piece.weight_lbs
        moment += slot.piece.weight_lbs * center

    if total_cargo == 0:
        return {}, []

    trailer_rxn = moment / tandem_pos
    pin_weight  = total_cargo - trailer_rxn

    # 5th wheel is behind drives (overhang); adds slightly more than pin to drives,
    # reduces steer. Approximate: drives get 1.05× pin, steer loses ~0.05× pin.
    steer_weight = TRAILER["steer_tare"] - round(pin_weight * 0.05)
    steer_weight = max(steer_weight, TRAILER["steer_tare"] - 2500)
    drive_weight  = TRAILER["drive_tare"] + round(pin_weight * 1.05)
    truck_gross = steer_weight + drive_weight + round(trailer_rxn)

    aw = {
        "steer":        round(steer_weight),
        "drive":        round(drive_weight),
        "trailer":      round(trailer_rxn),
        "freight_gross": round(total_cargo),
        "pin":          round(pin_weight),
    }

    v = []
    if aw["steer"]   > TRAILER["steer_limit"]:
        v.append(f"STEER {aw['steer']:,} lb > {TRAILER['steer_limit']:,} lb limit")
    if aw["drive"]   > TRAILER["drive_limit"]:
        v.append(f"DRIVE {aw['drive']:,} lb > {TRAILER['drive_limit']:,} lb limit")
    if aw["trailer"] > TRAILER["trailer_limit"]:
        v.append(f"TRAILER TANDEM {aw['trailer']:,} lb > {TRAILER['trailer_limit']:,} lb limit")
    if truck_gross   > TRAILER["gross_limit"]:
        v.append(f"GROSS {truck_gross:,} lb > {TRAILER['gross_limit']:,} lb limit")

    return aw, v


# ── Save / Load plan files ────────────────────────────────────────────────────
def _safe_name(s: str) -> str:
    """Strip characters unsafe for filenames, collapse whitespace to underscores."""
    return re.sub(r'[^\w-]', '_', s).strip('_')

def make_base_filename(info: ShipmentInfo) -> str:
    """
    Build the standard base name: Customer_BOL_DATE
    Parts are omitted when blank. Spaces/special chars become underscores.
    """
    parts = [_safe_name(p) for p in [info.customer, info.bol_num, info.date] if p.strip()]
    return '_'.join(parts) if parts else 'load_plan'

def save_plan_json(pieces: list, info: ShipmentInfo, limits: dict,
                   excluded_ids: list, override: bool, path: str):
    data = {
        "version": 1,
        "saved_at": datetime.datetime.now().isoformat(timespec='seconds'),
        "shipment": {
            "customer":      info.customer,
            "shipment_name": info.shipment_name,
            "trailer_num":   info.trailer_num,
            "bol_num":       info.bol_num,
            "date":          info.date,
        },
        "axle_limits": limits,
        "override":     override,
        "excluded_ids": excluded_ids,
        "freight": [
            {
                "piece_id":    p.piece_id,
                "crate_label": p.crate_label,
                "part_number": p.part_number,
                "length_in":   p.length_in,
                "width_in":    p.width_in,
                "height_in":   p.height_in,
                "weight_lbs":  p.weight_lbs,
                "pcs":         p.pcs,
                "notes":       p.notes,
                "placement":   p.placement,
            }
            for p in pieces
        ],
    }
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)

def load_plan_json(path: str):
    """Returns (pieces, info, limits, excluded_ids, override)."""
    with open(path) as f:
        data = json.load(f)
    s = data['shipment']
    info = ShipmentInfo(
        customer=s.get('customer', ''),
        shipment_name=s.get('shipment_name', ''),
        trailer_num=s.get('trailer_num', ''),
        bol_num=s.get('bol_num', ''),
        date=s.get('date', ''),
    )
    pieces = [
        FreightPiece(
            piece_id=p['piece_id'],
            crate_label=p['crate_label'],
            part_number=p['part_number'],
            length_in=p['length_in'],
            width_in=p['width_in'],
            height_in=p['height_in'],
            weight_lbs=p['weight_lbs'],
            pcs=p.get('pcs', 0),
            notes=p.get('notes', ''),
            placement=p.get('placement', 'auto'),
        )
        for p in data['freight']
    ]
    limits = data.get('axle_limits', {'steer': 12000, 'drive': 34000, 'trailer': 34000})
    excluded_ids = data.get('excluded_ids', [])
    override = data.get('override', False)
    return pieces, info, limits, excluded_ids, override


# ── PDF helpers ───────────────────────────────────────────────────────────────
def _pdf_writer(path: str):
    w = QPdfWriter(path)
    w.setPageSize(QPageSize(QPageSize.PageSizeId.Letter))
    w.setPageOrientation(QPageLayout.Orientation.Portrait)
    w.setResolution(150)
    return w


# ── PDF: Load Plan Floor Plan ─────────────────────────────────────────────────
def export_load_plan_pdf(plan: LoadPlan, info: ShipmentInfo, path: str):
    writer = _pdf_writer(path)
    painter = QPainter(writer)
    PW = writer.width()
    PH = writer.height()
    _draw_floor_plan(painter, PW, PH, plan, info)
    painter.end()


def _draw_floor_plan(painter: QPainter, PW: int, PH: int,
                     plan: LoadPlan, info: ShipmentInfo):
    M = 72       # margin ~0.48"
    LM = M + 50  # extra left for row labels

    # ── Title ──────────────────────────────────────────────────────────
    painter.setPen(QColor("#111"))
    font = QFont("Arial", 20, QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(QRectF(LM, M, PW - LM - M, 48),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                     "Trailer Loading Floor Plan")

    font = QFont("Arial", 10)
    painter.setFont(font)
    # Deduplicate center pieces (two slots share one piece)
    seen = set()
    total_w, n_crates = 0, 0
    for s in plan.slots:
        if s.piece and s.piece.piece_id not in seen:
            seen.add(s.piece.piece_id)
            total_w += s.piece.weight_lbs
            n_crates += 1
    excl = (f"Crates {', '.join(str(x) for x in plan.excluded_ids)} excluded  |  "
            if plan.excluded_ids else "")
    sub = (f"{info.customer or info.shipment_name}  |  {TRAILER['name']}  |  "
           f"{excl}Total: {total_w:,.0f} lb / {n_crates} crates")
    painter.drawText(QRectF(LM, M + 50, PW - LM - M, 26),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter, sub)

    detail_parts = []
    if info.date:        detail_parts.append(f"Date: {info.date}")
    if info.trailer_num: detail_parts.append(f"Trailer: {info.trailer_num}")
    if info.bol_num:     detail_parts.append(f"BOL: {info.bol_num}")
    if detail_parts:
        painter.drawText(QRectF(LM, M + 76, PW - LM - M, 22),
                         Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                         "  |  ".join(detail_parts))

    # ── Column headers ──────────────────────────────────────────────────
    col_w = (PW - LM - M) / 2
    hdr_y = M + 108

    font = QFont("Arial", 12, QFont.Weight.Bold)
    painter.setFont(font)
    painter.setPen(QColor("#222"))
    painter.drawText(QRectF(LM, hdr_y, PW - LM - M, 28),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                     "NOSE  (load first)")

    hdr_y += 30
    font = QFont("Arial", 11, QFont.Weight.Bold)
    painter.setFont(font)
    painter.setPen(TEAL_DARK)
    painter.drawText(QRectF(LM, hdr_y, col_w, 24),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                     "Position A")
    painter.drawText(QRectF(LM + col_w, hdr_y, col_w, 24),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                     "Position B")

    # ── Grid ────────────────────────────────────────────────────────────
    grid_top = hdr_y + 28
    rows = sorted(set(s.row for s in plan.slots))
    n_rows = len(rows)

    TAIL_RESERVE = 100   # space for tail label + notes at bottom
    grid_h = PH - M - TAIL_RESERVE - grid_top
    row_h = grid_h / n_rows if n_rows else 40

    slot_map = {(s.row, s.col): s for s in plan.slots}

    font_rlbl = QFont("Arial", 9, QFont.Weight.Bold)
    font_big   = QFont("Arial", 11, QFont.Weight.Bold)
    font_small = QFont("Arial", 9)

    for r in rows:
        y = grid_top + r * row_h

        # Row label
        painter.setFont(font_rlbl)
        painter.setPen(QColor("#444"))
        painter.drawText(QRectF(M, y, 46, row_h),
                         Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                         f"R{r + 1}")

        for c in range(2):
            slot = slot_map.get((r, c))
            x = LM + c * col_w

            # Center pieces: skip col=1 (already drawn by col=0)
            if slot and slot.is_center and c == 1:
                continue

            if slot and slot.piece:
                if slot.is_center:
                    full_w = col_w * 2
                    crate_ratio = min(slot.piece.width_in / TRAILER["interior_width_in"], 1.0)
                    gap = (full_w * (1.0 - crate_ratio)) / 2
                    cell = QRectF(LM + gap + 2, y + 2, full_w * crate_ratio - 4, row_h - 4)
                else:
                    cell = QRectF(x + 2, y + 2, col_w - 4, row_h - 4)

                painter.setBrush(QBrush(TEAL_LIGHT))
                painter.setPen(QPen(TEAL_DARK, 1.5))
                painter.drawRoundedRect(cell, 5, 5)

                top_half = QRectF(cell.x() + 4, cell.y() + 2,
                                  cell.width() - 8, cell.height() * 0.55)
                bot_half = QRectF(cell.x() + 4, cell.y() + cell.height() * 0.55,
                                  cell.width() - 8, cell.height() * 0.42)

                painter.setFont(font_big)
                painter.setPen(TEAL_DARK)
                painter.drawText(top_half,
                                 Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                                 slot.piece.crate_label)
                painter.setFont(font_small)
                painter.drawText(bot_half,
                                 Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                                 f"{slot.piece.weight_lbs:,.0f} lb")
            else:
                cell = QRectF(x + 2, y + 2, col_w - 4, row_h - 4)
                painter.setBrush(QBrush(VOID_BG))
                pen = QPen(GRAY_LINE, 1, Qt.PenStyle.DashLine)
                painter.setPen(pen)
                painter.drawRoundedRect(cell, 5, 5)
                painter.setFont(font_small)
                painter.setPen(QColor("#999"))
                painter.drawText(cell, Qt.AlignmentFlag.AlignCenter,
                                 "Open\nBlock / void-fill")

    # ── Tail label ─────────────────────────────────────────────────────
    tail_y = PH - M - TAIL_RESERVE + 6
    font = QFont("Arial", 12, QFont.Weight.Bold)
    painter.setFont(font)
    painter.setPen(QColor("#111"))
    painter.drawText(QRectF(LM, tail_y, PW - LM - M, 28),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                     "TAIL  (rear doors, load last)")

    # Clear space note
    clear_in = TRAILER["interior_length_in"] - plan.total_length_used_in
    clear_ft = clear_in / 12.0
    font = QFont("Arial", 8)
    painter.setFont(font)
    painter.setPen(QColor("#555"))
    painter.drawText(QRectF(LM, tail_y + 30, PW - LM - M, 20),
                     Qt.AlignmentFlag.AlignHCenter,
                     f"~{clear_ft:.1f} ft clear space behind row {n_rows} for load bars / void-fill")

    note = ("Load strictly R1 to R13 in order.  "
            "Weigh loaded trailer on a certified scale and adjust tandem slide as needed before transport.")
    painter.drawText(QRectF(LM, tail_y + 50, PW - LM - M, 30),
                     Qt.AlignmentFlag.AlignHCenter | Qt.TextFlag.TextWordWrap, note)

    # Axle weight summary
    aw = plan.axle_weights
    if aw:
        aw_str = (f"Est. Axle Weights  —  "
                  f"Steer: {aw.get('steer', 0):,} lb  |  "
                  f"Drive: {aw.get('drive', 0):,} lb  |  "
                  f"Trailer: {aw.get('trailer', 0):,} lb  |  "
                  f"Freight Gross: {aw.get('freight_gross', 0):,} lb")
        painter.setPen(TEAL_DARK)
        painter.drawText(QRectF(LM, PH - M - 16, PW - LM - M, 16),
                         Qt.AlignmentFlag.AlignHCenter, aw_str)


# ── PDF: Crate Labels ─────────────────────────────────────────────────────────
def export_crate_labels_pdf(plan: LoadPlan, info: ShipmentInfo, path: str):
    filled = sorted([s for s in plan.slots if s.piece], key=lambda s: (s.row, s.col))
    if not filled:
        return

    writer = _pdf_writer(path)
    painter = QPainter(writer)
    PW = writer.width()
    PH = writer.height()

    for i, slot in enumerate(filled):
        if i > 0:
            writer.newPage()
        _draw_crate_label(painter, PW, PH, slot, plan, info)

    painter.end()


def _draw_crate_label(painter: QPainter, PW: int, PH: int,
                      slot: PalletSlot, plan: LoadPlan, info: ShipmentInfo):
    PAD = 18    # outer border inset
    INNER_X = PAD + 22
    INNER_W = PW - 2 * (PAD + 22)

    # Outer border (tan/gold)
    painter.setPen(QPen(TAN, 1.5))
    painter.setBrush(Qt.BrushStyle.NoBrush)
    painter.drawRect(QRectF(PAD, PAD, PW - 2 * PAD, PH - 2 * PAD))

    # ── Top band: shipment name ─────────────────────────────────────────
    sep1_y = PH * 0.16

    font = QFont("Arial", 11)
    painter.setFont(font)
    painter.setPen(QColor("#555"))
    hdr = info.customer or info.shipment_name or "Freight Shipment"
    if info.date:
        hdr += f"   |   {info.date}"
    painter.drawText(QRectF(INNER_X, PAD + 10, INNER_W, sep1_y - PAD - 14),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter, hdr)

    painter.setPen(QPen(GRAY_LINE, 1))
    painter.drawLine(QPointF(INNER_X, sep1_y), QPointF(INNER_X + INNER_W, sep1_y))

    # ── Teal main box ──────────────────────────────────────────────────
    box_top = sep1_y + 18
    box_h   = PH * 0.30
    box_x   = INNER_X + 10
    box_w   = INNER_W - 20
    box     = QRectF(box_x, box_top, box_w, box_h)

    painter.setBrush(QBrush(TEAL_LIGHT))
    painter.setPen(QPen(TEAL_DARK, 2.5))
    painter.drawRoundedRect(box, 16, 16)

    # Crate number — very large
    font = QFont("Arial", 54, QFont.Weight.Bold)
    painter.setFont(font)
    painter.setPen(TEAL_DARK)
    painter.drawText(QRectF(box_x + 10, box_top + 16, box_w - 20, box_h * 0.58),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                     slot.piece.crate_label.upper())

    # Position label
    side_letter = "A" if slot.col == 0 else "B"
    side_name   = "Left" if slot.col == 0 else "Right"
    pos_str = f"Position {side_letter}  ({side_name})   ·   Row {slot.row + 1}"
    font = QFont("Arial", 16)
    painter.setFont(font)
    painter.drawText(QRectF(box_x + 10, box_top + box_h * 0.62, box_w - 20, box_h * 0.34),
                     Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter, pos_str)

    # ── Middle separator ───────────────────────────────────────────────
    sep2_y = box_top + box_h + 22
    painter.setPen(QPen(GRAY_LINE, 1))
    painter.drawLine(QPointF(INNER_X, sep2_y), QPointF(INNER_X + INNER_W, sep2_y))

    # ── Detail section ─────────────────────────────────────────────────
    detail_lines = []
    if slot.piece.part_number:
        detail_lines.append(("Part Number:", slot.piece.part_number))
    detail_lines.append(("Dimensions:",
                          f'{slot.piece.length_in:.0f}" × {slot.piece.width_in:.0f}" × {slot.piece.height_in:.0f}"  (L × W × H)'))
    detail_lines.append(("Weight:", f"{slot.piece.weight_lbs:,.0f} lbs"))
    if slot.piece.pcs:
        detail_lines.append(("Pieces:", f"{slot.piece.pcs:,}"))
    if slot.piece.notes:
        detail_lines.append(("Notes:", slot.piece.notes))

    LBL_W = 145
    ROW_H = 40
    det_y = sep2_y + 14
    font_lbl = QFont("Arial", 11)
    font_val = QFont("Arial", 13, QFont.Weight.Bold)

    for label, val in detail_lines:
        painter.setFont(font_lbl)
        painter.setPen(QColor("#666"))
        painter.drawText(QRectF(INNER_X + 20, det_y, LBL_W, ROW_H - 4),
                         Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, label)
        painter.setFont(font_val)
        painter.setPen(QColor("#111"))
        painter.drawText(QRectF(INNER_X + LBL_W + 10, det_y, INNER_W - LBL_W - 20, ROW_H - 4),
                         Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, val)
        det_y += ROW_H

    # ── Lower separator ────────────────────────────────────────────────
    sep3_y = det_y + 10
    painter.setPen(QPen(GRAY_LINE, 1))
    painter.drawLine(QPointF(INNER_X, sep3_y), QPointF(INNER_X + INNER_W, sep3_y))

    # ── Shipment footer ────────────────────────────────────────────────
    footer_lines = []
    if info.customer:        footer_lines.append(("Customer:", info.customer))
    if info.shipment_name:   footer_lines.append(("Shipment:", info.shipment_name))
    if info.bol_num:         footer_lines.append(("BOL #:", info.bol_num))
    if info.trailer_num:     footer_lines.append(("Trailer #:", info.trailer_num))

    fot_y = sep3_y + 14
    font_lbl2 = QFont("Arial", 11)
    font_val2 = QFont("Arial", 11)

    for label, val in footer_lines:
        painter.setFont(font_lbl2)
        painter.setPen(QColor("#666"))
        painter.drawText(QRectF(INNER_X + 20, fot_y, LBL_W, 30),
                         Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, label)
        painter.setFont(font_val2)
        painter.setPen(QColor("#222"))
        painter.drawText(QRectF(INNER_X + LBL_W + 10, fot_y, INNER_W - LBL_W - 20, 30),
                         Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, val)
        fot_y += 32


# ── Live Trailer View widget ──────────────────────────────────────────────────
class TrailerView(QWidget):
    plan_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.plan: Optional[LoadPlan] = None
        self.setMinimumSize(320, 480)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._drag_src: Optional[tuple] = None
        self._hover_slot: Optional[tuple] = None
        self.setMouseTracking(True)

    def set_plan(self, plan: LoadPlan):
        self.plan = plan
        self._drag_src = None
        self._hover_slot = None
        self.update()

    # ── Layout constants (shared between hit-testing and drawing) ────────────
    def _layout(self):
        M, NOSE_H, ROW_LBL = 28, 22, 36
        W, H = self.width(), self.height()
        grid_x = M + ROW_LBL
        grid_w = W - grid_x - M
        grid_y = M + NOSE_H + 4
        grid_h = H - grid_y - M - 20
        col_w  = grid_w / 2
        return grid_x, grid_w, grid_y, grid_h, col_w, M, NOSE_H, ROW_LBL

    # ── Hit-test: which slot is under a mouse position ───────────────────────
    def _slot_at(self, pos) -> Optional[tuple]:
        if not self.plan or not self.plan.slots:
            return None
        grid_x, grid_w, grid_y, grid_h, col_w, *_ = self._layout()
        trailer_len = TRAILER["interior_length_in"]
        px, py = pos.x(), pos.y()
        if not (grid_x <= px <= grid_x + grid_w and grid_y <= py <= grid_y + grid_h):
            return None
        col = 0 if px < grid_x + col_w else 1
        slot_map = {(s.row, s.col): s for s in self.plan.slots}
        for r in sorted(set(s.row for s in self.plan.slots)):
            sa = slot_map.get((r, 0))
            if sa and sa.piece:
                y0 = grid_y + (sa.row_start_in / trailer_len) * grid_h
                y1 = y0 + (sa.piece.length_in / trailer_len) * grid_h
                if y0 <= py < y1:
                    return (r, 0) if sa.is_center else (r, col)
        return None

    # ── Swap two pieces between slots ────────────────────────────────────────
    def _swap_pieces(self, src: tuple, dst: tuple):
        slot_map = {(s.row, s.col): s for s in self.plan.slots}
        s1, s2 = slot_map.get(src), slot_map.get(dst)
        if not s1 or s2 is None or s1.is_center or s2.is_center:
            return
        s1.piece, s2.piece = s2.piece, s1.piece
        self.plan.axle_weights, self.plan.violations = _calc_axle_weights(
            self.plan.slots, self.plan.override)
        self.plan_changed.emit()
        self.update()

    # ── Mouse events ─────────────────────────────────────────────────────────
    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        key = self._slot_at(event.pos())
        if key:
            slot_map = {(s.row, s.col): s for s in self.plan.slots}
            s = slot_map.get(key)
            if s and s.piece and not s.is_center:
                self._drag_src = key
                self.setCursor(Qt.CursorShape.ClosedHandCursor)

    def mouseMoveEvent(self, event):
        key = self._slot_at(event.pos())
        if key != self._hover_slot:
            self._hover_slot = key
            self.update()
        if self._drag_src:
            slot_map = {(s.row, s.col): s for s in self.plan.slots}
            tgt = slot_map.get(key) if key else None
            can_drop = (key and key != self._drag_src and
                        tgt is not None and not (tgt.is_center and tgt.piece))
            self.setCursor(Qt.CursorShape.DragMoveCursor if can_drop
                           else Qt.CursorShape.ClosedHandCursor)
        else:
            slot_map = {(s.row, s.col): s for s in self.plan.slots} if self.plan else {}
            s = slot_map.get(key) if key else None
            if s and s.piece and not s.is_center:
                self.setCursor(Qt.CursorShape.OpenHandCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._drag_src:
            dst = self._slot_at(event.pos())
            if dst and dst != self._drag_src:
                self._swap_pieces(self._drag_src, dst)
            self._drag_src = None
            self.setCursor(Qt.CursorShape.ArrowCursor)
            self.update()

    def leaveEvent(self, event):
        self._drag_src = None
        self._hover_slot = None
        self.setCursor(Qt.CursorShape.ArrowCursor)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        self._draw(painter)

    def _draw(self, painter: QPainter):
        if not self.plan or not self.plan.slots:
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                             "No load plan — press Calculate.")
            return

        slots  = self.plan.slots
        rows   = sorted(set(s.row for s in slots))
        n_rows = len(rows)

        grid_x, grid_w, grid_y, grid_h, col_w, M, NOSE_H, ROW_LBL = self._layout()
        H = self.height()
        trailer_len = TRAILER["interior_length_in"]
        interior_w  = TRAILER["interior_width_in"]

        # Trailer outline
        painter.setPen(QPen(QColor("#444"), 2))
        painter.setBrush(QBrush(QColor("#f5f5f0")))
        painter.drawRect(QRectF(grid_x, grid_y, grid_w, grid_h))

        # Nose label
        font = QFont("Arial", 8, QFont.Weight.Bold)
        painter.setFont(font)
        painter.setPen(QColor("#333"))
        painter.drawText(QRectF(grid_x, M, grid_w, NOSE_H),
                         Qt.AlignmentFlag.AlignCenter, "◀  NOSE")

        # Axle reference lines
        tandem_in = TRAILER["tandem_from_nose_in"]
        for depth, label, color in [
            (36,        "Drive Axles",   "#e67e00"),
            (tandem_in, "Trailer Axles", "#c0392b"),
        ]:
            ay = grid_y + (depth / trailer_len) * grid_h
            painter.setPen(QPen(QColor(color), 1.5, Qt.PenStyle.DashLine))
            painter.drawLine(QPointF(grid_x, ay), QPointF(grid_x + grid_w, ay))
            painter.setPen(QColor(color))
            painter.setFont(QFont("Arial", 6))
            painter.drawText(QRectF(grid_x + 2, ay - 12, grid_w - 4, 12),
                             Qt.AlignmentFlag.AlignRight, label)

        COLORS = [
            "#4a90d9","#e67e22","#27ae60","#8e44ad","#c0392b","#16a085",
            "#d35400","#2980b9","#7f8c8d","#f39c12","#1abc9c","#e74c3c",
        ]
        slot_map   = {(s.row, s.col): s for s in slots}
        font_big   = QFont("Arial", 7, QFont.Weight.Bold)
        font_small = QFont("Arial", 5)
        font_rlbl  = QFont("Arial", 7, QFont.Weight.Bold)

        for r in rows:
            sa = slot_map.get((r, 0))
            if not sa:
                continue
            y_px     = grid_y + (sa.row_start_in / trailer_len) * grid_h
            piece_h  = ((sa.piece.length_in / trailer_len) * grid_h) if sa.piece else (grid_h / n_rows)

            # Row label
            painter.setFont(font_rlbl)
            painter.setPen(QColor("#444"))
            painter.drawText(QRectF(M, y_px, ROW_LBL - 2, piece_h),
                             Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                             f"R{r+1}")

            for c in range(2):
                slot = slot_map.get((r, c))
                x = grid_x + c * col_w

                if slot and slot.piece:
                    if slot.col == 1 and slot.is_center:
                        continue  # center piece drawn once via col=0
                    ph = (slot.piece.length_in / trailer_len) * grid_h
                    if slot.is_center:
                        # Actual width centered with gaps on each side
                        crate_ratio = min(slot.piece.width_in / interior_w, 1.0)
                        gap_px = (grid_w * (1.0 - crate_ratio)) / 2
                        cell = QRectF(grid_x + gap_px + 1, y_px + 1,
                                      grid_w * crate_ratio - 2, ph - 2)
                    else:
                        cell = QRectF(x + 1, y_px + 1, col_w - 2, ph - 2)

                    color = QColor(COLORS[(slot.piece.piece_id - 1) % len(COLORS)])
                    painter.setBrush(QBrush(color))
                    painter.setPen(QPen(QColor("#222"), 0.5))
                    painter.drawRect(cell)

                    painter.setFont(font_big)
                    painter.setPen(Qt.GlobalColor.white)
                    if slot.is_center:
                        painter.drawText(cell.adjusted(2, 2, -2, -2),
                                         Qt.AlignmentFlag.AlignCenter,
                                         f"CENTER\n{slot.piece.crate_label}\n{slot.piece.weight_lbs:,.0f} lb")
                    else:
                        painter.drawText(cell.adjusted(2, 2, -2, -cell.height() // 2),
                                         Qt.AlignmentFlag.AlignCenter, slot.piece.crate_label)
                        painter.setFont(font_small)
                        painter.drawText(cell.adjusted(2, cell.height() // 2, -2, -2),
                                         Qt.AlignmentFlag.AlignCenter,
                                         f"{slot.piece.weight_lbs:,.0f} lb")

                    # Drag / drop highlights
                    norm_key = (r, 0) if slot.is_center else (r, c)
                    if self._drag_src and norm_key == self._drag_src:
                        painter.setBrush(Qt.BrushStyle.NoBrush)
                        painter.setPen(QPen(QColor("#f39c12"), 2.5))
                        painter.drawRect(cell)
                    elif (self._drag_src and self._hover_slot and
                          norm_key == self._hover_slot and norm_key != self._drag_src):
                        painter.setBrush(Qt.BrushStyle.NoBrush)
                        painter.setPen(QPen(QColor("#27ae60"), 2.5))
                        painter.drawRect(cell)
                else:
                    cell = QRectF(x + 1, y_px + 1, col_w - 2, piece_h - 2)
                    painter.setBrush(QBrush(QColor("#e8e8e8")))
                    painter.setPen(QPen(QColor("#bbb"), 0.5, Qt.PenStyle.DashLine))
                    painter.drawRect(cell)
                    # Empty slot as drop target
                    if (self._drag_src and self._hover_slot and
                            (r, c) == self._hover_slot and (r, c) != self._drag_src):
                        painter.setBrush(Qt.BrushStyle.NoBrush)
                        painter.setPen(QPen(QColor("#27ae60"), 2.5))
                        painter.drawRect(cell)

        # Column labels at bottom
        painter.setFont(QFont("Arial", 7))
        painter.setPen(QColor("#555"))
        for c, lbl in enumerate(["Position A (Left)", "Position B (Right)"]):
            painter.drawText(QRectF(grid_x + c * col_w, H - 20, col_w, 18),
                             Qt.AlignmentFlag.AlignCenter, lbl)


# ── Excel Import Dialog ───────────────────────────────────────────────────────
class ImportDialog(QDialog):
    """Preview packing list from Excel, check/uncheck rows to include/exclude."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import Packing List from Excel")
        self.setMinimumSize(820, 520)
        self.pieces: list[FreightPiece] = []
        self.excluded_ids: list[int] = []

        layout = QVBoxLayout(self)

        # File picker
        file_row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Select Excel packing list (.xlsx)…")
        self.path_edit.setReadOnly(True)
        btn_browse = QPushButton("Browse…")
        btn_browse.clicked.connect(self._browse)
        file_row.addWidget(self.path_edit)
        file_row.addWidget(btn_browse)
        layout.addLayout(file_row)

        lbl = QLabel("Check rows to include. Uncheck to exclude from the load plan.")
        lbl.setStyleSheet("color:#555; font-size:11px;")
        layout.addWidget(lbl)

        # Preview table
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(
            ["✓ Include", "Crate #", "Part Number", "L\"", "W\"", "H\"", "Weight (lbs)", "PCS", "Notes"])
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self._accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Packing List", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.path_edit.setText(path)
            self._load(path)

    def _load(self, path: str):
        if not HAS_OPENPYXL:
            QMessageBox.critical(self, "Missing Library",
                                 "openpyxl is required for Excel import.\n"
                                 "Run: pip3 install openpyxl")
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            self.table.setRowCount(0)
            self._raw_rows = []

            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] is None:
                    continue
                try:
                    crate_id = int(row[0])
                except (TypeError, ValueError):
                    continue

                part   = str(row[1] or "")
                l_in   = float(row[3] or 0)
                w_in   = float(row[4] or 0)
                h_in   = float(row[5] or 0)
                weight = float(row[6] or 0)
                pcs    = int(row[7] or 0)
                notes  = str(row[9] or "")

                self._raw_rows.append((crate_id, part, l_in, w_in, h_in, weight, pcs, notes))

                r = self.table.rowCount()
                self.table.insertRow(r)

                chk = QCheckBox()
                chk.setChecked(True)
                chk_widget = QWidget()
                chk_layout = QHBoxLayout(chk_widget)
                chk_layout.addWidget(chk)
                chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
                chk_layout.setContentsMargins(0, 0, 0, 0)
                self.table.setCellWidget(r, 0, chk_widget)

                for col, val in enumerate([crate_id, part, l_in, w_in, h_in,
                                           f"{weight:,.0f}", pcs, notes], start=1):
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.table.setItem(r, col, item)

        except Exception as e:
            QMessageBox.critical(self, "Import Error", str(e))

    def _accept(self):
        self.pieces = []
        self.excluded_ids = []

        for r, raw in enumerate(self._raw_rows):
            crate_id, part, l_in, w_in, h_in, weight, pcs, notes = raw
            chk_widget = self.table.cellWidget(r, 0)
            chk = chk_widget.findChild(QCheckBox)
            if chk and chk.isChecked():
                self.pieces.append(FreightPiece(
                    piece_id=crate_id,
                    crate_label=f"Crate {crate_id}",
                    part_number=part,
                    length_in=l_in,
                    width_in=w_in,
                    height_in=h_in,
                    weight_lbs=weight,
                    pcs=pcs,
                    notes=notes,
                ))
            else:
                self.excluded_ids.append(crate_id)

        self.accept()


# ── Main Window ───────────────────────────────────────────────────────────────
class FreightLoaderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Freight Loader — 53' Dry Van")
        self.setMinimumSize(1150, 760)
        self._next_id = 1
        self._plan: Optional[LoadPlan] = None
        self._excluded_ids: list = []
        self._current_file: Optional[str] = None   # path of the open .json plan
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setSpacing(6)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(splitter)

        # ══ LEFT PANEL ══════════════════════════════════════════════════
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setSpacing(6)

        # ── File actions ───────────────────────────────────────────────
        file_row = QHBoxLayout()
        btn_new   = QPushButton("New Plan")
        btn_open  = QPushButton("Open Plan…")
        btn_save  = QPushButton("Save Plan")
        btn_saveas = QPushButton("Save As…")
        for b in [btn_new, btn_open, btn_save, btn_saveas]:
            b.setFixedHeight(30)
            file_row.addWidget(b)
        btn_new.clicked.connect(self._new_plan)
        btn_open.clicked.connect(self._open_plan)
        btn_save.clicked.connect(self._save_plan)
        btn_saveas.clicked.connect(self._save_plan_as)
        ll.addLayout(file_row)

        # ── Shipment Info ──────────────────────────────────────────────
        grp_info = QGroupBox("Shipment Info")
        info_form = QFormLayout(grp_info)
        info_form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.DontWrapRows)

        self.f_customer      = QLineEdit(); self.f_customer.setPlaceholderText("Customer name")
        self.f_shipment_name = QLineEdit(); self.f_shipment_name.setPlaceholderText("Shipment / load name")
        self.f_trailer_num   = QLineEdit(); self.f_trailer_num.setPlaceholderText("Optional")
        self.f_bol           = QLineEdit(); self.f_bol.setPlaceholderText("Optional")
        self.f_date          = QLineEdit(datetime.date.today().strftime("%Y-%m-%d"))

        info_form.addRow("Customer:",      self.f_customer)
        info_form.addRow("Shipment Name:", self.f_shipment_name)
        info_form.addRow("Trailer #:",     self.f_trailer_num)
        info_form.addRow("BOL #:",         self.f_bol)
        info_form.addRow("Date:",          self.f_date)
        ll.addWidget(grp_info)

        # ── Freight table ──────────────────────────────────────────────
        grp_freight = QGroupBox("Freight Pieces")
        gl = QVBoxLayout(grp_freight)

        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(
            ["Crate #", "Label", "Part #", 'L"', 'W"', 'H"', "Weight (lbs)", "PCS", "Pos."])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for col, w in [(0,58),(3,46),(4,46),(5,46),(6,88),(7,50),(8,72)]:
            self.table.setColumnWidth(col, w)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        gl.addWidget(self.table)

        btn_row = QHBoxLayout()
        btn_import = QPushButton("Import from Excel…")
        btn_import.clicked.connect(self._import_excel)
        btn_add    = QPushButton("+ Add Row")
        btn_add.clicked.connect(self._add_row)
        btn_del    = QPushButton("Remove Selected")
        btn_del.clicked.connect(self._remove_row)
        btn_clear  = QPushButton("Clear All")
        btn_clear.clicked.connect(self._clear_table)
        for b in [btn_import, btn_add, btn_del, btn_clear]:
            btn_row.addWidget(b)
        gl.addLayout(btn_row)
        ll.addWidget(grp_freight)

        # ── Axle limits ────────────────────────────────────────────────
        grp_limits = QGroupBox("DOT Axle Weight Limits (lbs)")
        lim_row = QHBoxLayout(grp_limits)
        self._limit_fields = {}
        for key, label, default in [
            ("steer",   "Steer",   12000),
            ("drive",   "Drive",   34000),
            ("trailer", "Trailer", 34000),
        ]:
            vb = QVBoxLayout()
            vb.addWidget(QLabel(label, alignment=Qt.AlignmentFlag.AlignCenter))
            spin = QSpinBox()
            spin.setRange(1000, 200000)
            spin.setValue(default)
            spin.setSingleStep(500)
            spin.setGroupSeparatorShown(True)
            self._limit_fields[key] = spin
            vb.addWidget(spin)
            lim_row.addLayout(vb)
        self.chk_override = QCheckBox("Allow override")
        lim_row.addWidget(self.chk_override)
        ll.addWidget(grp_limits)

        # ── Calculate ─────────────────────────────────────────────────
        btn_calc = QPushButton("⚡  Calculate Load Plan")
        btn_calc.setFixedHeight(44)
        font = btn_calc.font(); font.setPointSize(13); font.setBold(True)
        btn_calc.setFont(font)
        btn_calc.clicked.connect(self._calculate)
        ll.addWidget(btn_calc)

        # ── Summary / violations ───────────────────────────────────────
        grp_sum = QGroupBox("Axle Weight Summary")
        sl = QVBoxLayout(grp_sum)
        self.lbl_summary = QLabel("—")
        self.lbl_summary.setTextFormat(Qt.TextFormat.RichText)
        self.lbl_summary.setWordWrap(True)
        sl.addWidget(self.lbl_summary)
        ll.addWidget(grp_sum)

        grp_status = QGroupBox("Status / Violations")
        stl = QVBoxLayout(grp_status)
        self.lbl_status = QLabel("No plan calculated.")
        self.lbl_status.setTextFormat(Qt.TextFormat.RichText)
        self.lbl_status.setWordWrap(True)
        stl.addWidget(self.lbl_status)
        ll.addWidget(grp_status)

        # ── Export buttons ─────────────────────────────────────────────
        exp_row = QHBoxLayout()
        self.btn_plan_pdf  = QPushButton("Export Floor Plan PDF")
        self.btn_label_pdf = QPushButton("Export Crate Labels PDF")
        for b in [self.btn_plan_pdf, self.btn_label_pdf]:
            b.setEnabled(False)
            exp_row.addWidget(b)
        self.btn_plan_pdf.clicked.connect(self._export_floor_plan)
        self.btn_label_pdf.clicked.connect(self._export_labels)
        ll.addLayout(exp_row)

        splitter.addWidget(left)

        # ══ RIGHT PANEL — trailer view ════════════════════════════════
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        lbl_view = QLabel("Trailer View — Top-Down (Nose at Top)")
        lbl_view.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_view.setStyleSheet("font-weight:bold; font-size:11px;")
        rl.addWidget(lbl_view)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.trailer_view = TrailerView()
        self.trailer_view.setMinimumHeight(600)
        self.trailer_view.plan_changed.connect(self._on_plan_changed_by_drag)
        scroll.setWidget(self.trailer_view)
        rl.addWidget(scroll)
        splitter.addWidget(right)
        splitter.setSizes([500, 610])

    def _on_plan_changed_by_drag(self):
        if self._plan:
            self._update_summary(self._plan)

    # ── Actions ──────────────────────────────────────────────────────────────
    def _get_shipment_info(self) -> ShipmentInfo:
        return ShipmentInfo(
            customer=self.f_customer.text().strip(),
            shipment_name=self.f_shipment_name.text().strip(),
            trailer_num=self.f_trailer_num.text().strip(),
            bol_num=self.f_bol.text().strip(),
            date=self.f_date.text().strip(),
        )

    def _import_excel(self):
        dlg = ImportDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.pieces:
            self._load_pieces(dlg.pieces)
            self._excluded_ids = dlg.excluded_ids
            if dlg.excluded_ids:
                self.lbl_status.setText(
                    f"Imported {len(dlg.pieces)} crates. "
                    f"Excluded: {', '.join('Crate ' + str(i) for i in dlg.excluded_ids)}")

    def _load_pieces(self, pieces: list):
        self.table.setRowCount(0)
        self._next_id = 1
        for p in pieces:
            self._add_row(
                crate_id=p.piece_id,
                label=p.crate_label,
                part_num=p.part_number,
                length=p.length_in,
                width=p.width_in,
                height=p.height_in,
                weight=p.weight_lbs,
                pcs=p.pcs,
                placement=p.placement,
            )

    def _add_row(self, crate_id=None, label="", part_num="",
                 length=48.0, width=48.0, height=48.0, weight=1000.0, pcs=0,
                 placement="auto"):
        r = self.table.rowCount()
        self.table.insertRow(r)

        if crate_id is None:
            crate_id = self._next_id
        if not label:
            label = f"Crate {crate_id}"

        self._next_id = max(self._next_id, crate_id) + 1

        for col, val in enumerate([crate_id, label, part_num,
                                    length, width, height, weight, pcs]):
            item = QTableWidgetItem(str(val))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if col == 0:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, col, item)

        # Col 8: Pos. placement combo box — changing it auto-recalculates
        combo = QComboBox()
        combo.addItems(["Auto", "Center"])
        combo.setCurrentIndex(1 if placement == "center" else 0)
        combo.currentIndexChanged.connect(lambda: self._calculate(silent=True))
        self.table.setCellWidget(r, 8, combo)

    def _remove_row(self):
        rows = sorted(set(i.row() for i in self.table.selectedItems()), reverse=True)
        for r in rows:
            self.table.removeRow(r)

    def _clear_table(self):
        if QMessageBox.question(self, "Clear All", "Remove all freight rows?",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                                ) == QMessageBox.StandardButton.Yes:
            self.table.setRowCount(0)
            self._next_id = 1

    def _read_pieces(self) -> list:
        pieces = []
        for r in range(self.table.rowCount()):
            try:
                crate_id = int(self.table.item(r, 0).text())
                label    = self.table.item(r, 1).text() if self.table.item(r, 1) else f"Crate {crate_id}"
                part     = self.table.item(r, 2).text() if self.table.item(r, 2) else ""
                l_in     = float(self.table.item(r, 3).text())
                w_in     = float(self.table.item(r, 4).text())
                h_in     = float(self.table.item(r, 5).text())
                weight   = float(self.table.item(r, 6).text())
                pcs_item = self.table.item(r, 7)
                pcs      = int(pcs_item.text()) if pcs_item and pcs_item.text().strip() else 0
                combo    = self.table.cellWidget(r, 8)
                placement = "center" if combo and combo.currentText() == "Center" else "auto"
                pieces.append(FreightPiece(crate_id, label, part, l_in, w_in, h_in, weight, pcs,
                                           placement=placement))
            except (ValueError, AttributeError) as e:
                QMessageBox.warning(self, "Input Error", f"Row {r+1}: {e}")
                return []
        return pieces

    def _calculate(self, silent: bool = False):
        pieces = self._read_pieces()
        if not pieces:
            return

        TRAILER["steer_limit"]   = self._limit_fields["steer"].value()
        TRAILER["drive_limit"]   = self._limit_fields["drive"].value()
        TRAILER["trailer_limit"] = self._limit_fields["trailer"].value()

        override = self.chk_override.isChecked()
        plan = plan_load(pieces, override=override, excluded_ids=self._excluded_ids)

        if plan.violations and not override and not silent:
            msg = "Weight violations detected:\n\n" + "\n".join(plan.violations)
            msg += "\n\nCheck 'Allow override' to proceed anyway, or adjust the load."
            QMessageBox.warning(self, "Weight Violations", msg)

        self._plan = plan
        self.trailer_view.set_plan(plan)
        self._update_summary(plan)
        self.btn_plan_pdf.setEnabled(True)
        self.btn_label_pdf.setEnabled(True)

    def _update_summary(self, plan: LoadPlan):
        aw = plan.axle_weights
        rows_html = ""
        for key, label, lim_key in [
            ("steer",   "Steer",          "steer_limit"),
            ("drive",   "Drive",          "drive_limit"),
            ("trailer", "Trailer Tandem", "trailer_limit"),
        ]:
            val   = aw.get(key, 0)
            limit = TRAILER[lim_key]
            pct   = val / limit * 100 if limit else 0
            color = "#c0392b" if val > limit else "#27ae60"
            rows_html += (f"<tr>"
                          f"<td><b>{label}:</b></td>"
                          f"<td align='right'><span style='color:{color}'>{val:,} lbs</span></td>"
                          f"<td align='right'>&nbsp;{pct:.0f}% of {limit:,}</td>"
                          f"</tr>")
        # Freight gross
        fg = aw.get("freight_gross", 0)
        rows_html += (f"<tr><td><b>Freight Gross:</b></td>"
                      f"<td align='right'>{fg:,} lbs</td><td></td></tr>")

        # Side weights
        left_w  = sum((s.piece.weight_lbs / 2 if s.is_center else s.piece.weight_lbs)
                       for s in plan.slots if s.piece and s.col == 0)
        right_w = sum((s.piece.weight_lbs / 2 if s.is_center else s.piece.weight_lbs)
                       for s in plan.slots if s.piece and s.col == 1)
        diff    = abs(left_w - right_w)
        heavier = "Left" if left_w > right_w else "Right"
        diff_color = "#c0392b" if diff > 2000 else "#e67e22" if diff > 500 else "#27ae60"
        rows_html += (
            f"<tr><td colspan='3'><hr style='margin:2px'></td></tr>"
            f"<tr><td><b>Left (A):</b></td><td align='right'>{left_w:,} lbs</td><td></td></tr>"
            f"<tr><td><b>Right (B):</b></td><td align='right'>{right_w:,} lbs</td><td></td></tr>"
            f"<tr><td><b>Side diff:</b></td>"
            f"<td align='right'><span style='color:{diff_color}'>{diff:,} lbs</span></td>"
            f"<td align='right'><span style='color:{diff_color}'>{heavier} heavier</span></td>"
            f"</tr>"
        )
        self.lbl_summary.setText(f"<table>{rows_html}</table>")

        if plan.violations:
            v_html = "<br>".join(f"<span style='color:red'>⚠ {v}</span>" for v in plan.violations)
            if plan.override:
                v_html += "<br><i>Override active.</i>"
            self.lbl_status.setText(v_html)
        else:
            self.lbl_status.setText(
                "<span style='color:green'>✓ All axle weights within legal limits.</span>")

    # ── File: New / Open / Save ───────────────────────────────────────────────
    def _new_plan(self):
        if QMessageBox.question(self, "New Plan", "Clear everything and start a new plan?",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                                ) != QMessageBox.StandardButton.Yes:
            return
        self.table.setRowCount(0)
        self._next_id = 1
        self._plan = None
        self._excluded_ids = []
        self._current_file = None
        for f in [self.f_customer, self.f_shipment_name, self.f_trailer_num, self.f_bol]:
            f.clear()
        self.f_date.setText(datetime.date.today().strftime("%Y-%m-%d"))
        for key, default in [("steer", 12000), ("drive", 34000), ("trailer", 34000)]:
            self._limit_fields[key].setValue(default)
        self.chk_override.setChecked(False)
        self.lbl_summary.setText("—")
        self.lbl_status.setText("No plan calculated.")
        self.trailer_view.set_plan(None)
        self.btn_plan_pdf.setEnabled(False)
        self.btn_label_pdf.setEnabled(False)
        self._set_title()

    def _open_plan(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Load Plan", SAVE_DIR, "Load Plan Files (*.json)")
        if not path:
            return
        try:
            pieces, info, limits, excluded_ids, override = load_plan_json(path)
        except Exception as e:
            QMessageBox.critical(self, "Open Failed", str(e))
            return

        # Populate shipment info
        self.f_customer.setText(info.customer)
        self.f_shipment_name.setText(info.shipment_name)
        self.f_trailer_num.setText(info.trailer_num)
        self.f_bol.setText(info.bol_num)
        self.f_date.setText(info.date)

        # Populate axle limits
        for key in ("steer", "drive", "trailer"):
            if key in limits:
                self._limit_fields[key].setValue(limits[key])
        self.chk_override.setChecked(override)

        # Populate freight table
        self._load_pieces(pieces)
        self._excluded_ids = excluded_ids

        # Recalculate and display
        self._current_file = path
        self._set_title()
        self._calculate()

    def _save_plan(self):
        if self._current_file:
            self._write_plan(self._current_file)
        else:
            self._save_plan_as()

    def _save_plan_as(self):
        info = self._get_shipment_info()
        default_name = make_base_filename(info) + ".json"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Load Plan", os.path.join(SAVE_DIR, default_name),
            "Load Plan Files (*.json)")
        if path:
            self._write_plan(path)
            self._current_file = path
            self._set_title()

    def _write_plan(self, path: str):
        pieces = self._read_pieces()
        if not pieces and self.table.rowCount() > 0:
            return   # parse error already shown by _read_pieces
        info = self._get_shipment_info()
        limits = {k: self._limit_fields[k].value() for k in ("steer", "drive", "trailer")}
        try:
            save_plan_json(pieces, info, limits, self._excluded_ids,
                           self.chk_override.isChecked(), path)
            self.lbl_status.setText(
                self.lbl_status.text() +
                f"<br><span style='color:#555'>Saved: {os.path.basename(path)}</span>")
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", str(e))

    def _set_title(self):
        name = os.path.basename(self._current_file) if self._current_file else "Unsaved"
        self.setWindowTitle(f"Freight Loader — {name}")

    # ── PDF exports ───────────────────────────────────────────────────────────
    def _export_floor_plan(self):
        if not self._plan:
            return
        info = self._get_shipment_info()
        default = make_base_filename(info) + "_load_plan.pdf"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Floor Plan PDF", os.path.join(SAVE_DIR, default),
            "PDF Files (*.pdf)")
        if path:
            export_load_plan_pdf(self._plan, info, path)
            QMessageBox.information(self, "Saved", f"Floor plan saved:\n{path}")

    def _export_labels(self):
        if not self._plan:
            return
        info = self._get_shipment_info()
        default = make_base_filename(info) + "_labels.pdf"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Crate Labels PDF", os.path.join(SAVE_DIR, default),
            "PDF Files (*.pdf)")
        if path:
            export_crate_labels_pdf(self._plan, info, path)
            QMessageBox.information(self, "Saved", f"Crate labels saved:\n{path}")


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = FreightLoaderApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
