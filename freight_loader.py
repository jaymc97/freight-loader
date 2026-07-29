import sys
import math
from dataclasses import dataclass, field
from typing import Optional
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLabel, QSpinBox,
    QDoubleSpinBox, QLineEdit, QGroupBox, QSplitter, QScrollArea,
    QMessageBox, QHeaderView, QCheckBox, QFrame, QSizePolicy,
    QDialog, QDialogButtonBox, QTextEdit, QFileDialog
)
from PyQt6.QtCore import Qt, QRectF, QPointF, QSizeF
from PyQt6.QtGui import (
    QPainter, QColor, QPen, QBrush, QFont, QPageLayout, QPageSize,
    QPdfWriter
)


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class FreightPiece:
    piece_id: int
    description: str
    length_in: float   # inches
    width_in: float
    height_in: float
    weight_lbs: float

@dataclass
class PalletSlot:
    """One slot in the trailer grid (row=distance from nose, col=0 left/1 right)."""
    row: int           # 0 = nose
    col: int           # 0 = left, 1 = right
    piece: Optional[FreightPiece] = None
    row_depth_in: float = 0.0   # cumulative depth from nose to start of this row

@dataclass
class LoadPlan:
    slots: list[PalletSlot]
    axle_weights: dict          # steer, drive, trailer
    violations: list[str]
    override: bool = False


# ---------------------------------------------------------------------------
# Trailer constants — 53' dry van
# ---------------------------------------------------------------------------

TRAILER = {
    "name": "53' Dry Van",
    "interior_length_in": 636,   # 53 ft
    "interior_width_in": 98,     # ~98" usable
    "interior_height_in": 110,
    # Axle positions measured from nose of trailer (kingpin ~= 0 ref)
    # These are approximate industry-standard numbers
    "kingpin_from_nose_in": 0,
    # Distance from kingpin to center of drive axles (~18" back of cab + ~~)
    # Typical: kingpin to tandem center = ~~36"
    # For weight calc we model the trailer as a beam on two supports:
    #   front support = drive axles at ~36" behind kingpin relative to trailer
    #   rear support  = trailer axles at rear
    # We convert everything to distance from nose of trailer.
    "drive_axle_from_nose_in": 36,    # kingpin area — drives bear weight here
    "trailer_axle_from_nose_in": 600, # ~50 ft from nose, at rear of 53'
    # Legal axle weight limits (lbs)
    "steer_limit": 12000,
    "drive_limit": 34000,
    "trailer_limit": 34000,
    "gross_limit": 80000,
    "tractor_steer_to_drive_in": 228,  # ~19 ft between steer and drive on tractor
}

COLS = 2        # pallet positions wide
SLOT_WIDTH_IN = TRAILER["interior_width_in"] / COLS   # ~49"


# ---------------------------------------------------------------------------
# Loading algorithm
# ---------------------------------------------------------------------------

def plan_load(pieces: list[FreightPiece], override: bool = False) -> LoadPlan:
    """
    Greedy best-fit: place pieces nose-to-tail, 2 wide per row.
    Heavier pieces go toward the drive axles (optimal DOT balance).
    Returns a LoadPlan with slot assignments and axle weight calculations.
    """
    if not pieces:
        return LoadPlan(slots=[], axle_weights={}, violations=[], override=override)

    # Sort: heaviest first so they land closest to drives
    sorted_pieces = sorted(pieces, key=lambda p: p.weight_lbs, reverse=True)

    slots: list[PalletSlot] = []
    cursor_depth = 0.0   # how far from nose we've consumed
    row = 0
    i = 0

    while i < len(sorted_pieces):
        # Grab up to 2 pieces for this row
        row_pieces = sorted_pieces[i:i + COLS]
        i += COLS

        # Row depth = max length of pieces in this row
        row_depth = max(p.length_in for p in row_pieces)

        if cursor_depth + row_depth > TRAILER["interior_length_in"]:
            # Doesn't fit — still assign (will flag as violation)
            pass

        for col, piece in enumerate(row_pieces):
            slot = PalletSlot(
                row=row,
                col=col,
                piece=piece,
                row_depth_in=cursor_depth,
            )
            slots.append(slot)

        cursor_depth += row_depth
        row += 1

    axle_weights, violations = _calc_axle_weights(slots, pieces, override)

    return LoadPlan(slots=slots, axle_weights=axle_weights,
                    violations=violations, override=override)


def _calc_axle_weights(slots, all_pieces, override):
    """
    Simple beam model: trailer supported at drive axles and trailer axles.
    Steer axle weight = (gross - drive_reaction) * steer_fraction (simplified).

    We calculate the reaction force at each support using lever/moment equations.
    """
    D = TRAILER["drive_axle_from_nose_in"]      # drive support position
    T = TRAILER["trailer_axle_from_nose_in"]    # trailer support position
    span = T - D  # distance between supports

    total_weight = sum(p.weight_lbs for p in all_pieces)

    # Moments about drive axle to find trailer axle reaction
    moment_about_drive = 0.0
    for slot in slots:
        if slot.piece is None:
            continue
        # Center of piece longitudinally
        piece_center = slot.row_depth_in + slot.piece.length_in / 2
        arm = piece_center - D   # + = toward tail, - = toward nose/steer
        moment_about_drive += slot.piece.weight_lbs * arm

    # Trailer axle reaction (upward)
    trailer_reaction = moment_about_drive / span if span > 0 else 0
    trailer_reaction = max(0, trailer_reaction)

    drive_reaction = total_weight - trailer_reaction

    # Steer axle takes a portion of drive reaction based on tractor geometry
    # Simplified: steer ~ 10-12% of gross for a typical loaded truck
    steer_to_drive = TRAILER["tractor_steer_to_drive_in"]
    # Moment of tractor weight is already embedded; use industry rule of thumb
    # steer = (drive_reaction) * (bogie_rear / wheelbase) — approx
    steer_frac = 0.12  # ~12% of gross on steers is typical
    steer_weight = total_weight * steer_frac
    drive_reaction = drive_reaction - steer_weight  # net on drives

    axle_weights = {
        "steer": round(steer_weight),
        "drive": round(drive_reaction),
        "trailer": round(trailer_reaction),
        "gross": round(total_weight + 20000),   # +20k tractor tare estimate
    }

    violations = []
    if axle_weights["steer"] > TRAILER["steer_limit"]:
        violations.append(f"STEER axle: {axle_weights['steer']:,} lbs exceeds {TRAILER['steer_limit']:,} lb limit")
    if axle_weights["drive"] > TRAILER["drive_limit"]:
        violations.append(f"DRIVE axles: {axle_weights['drive']:,} lbs exceeds {TRAILER['drive_limit']:,} lb limit")
    if axle_weights["trailer"] > TRAILER["trailer_limit"]:
        violations.append(f"TRAILER axles: {axle_weights['trailer']:,} lbs exceeds {TRAILER['trailer_limit']:,} lb limit")
    if axle_weights["gross"] > TRAILER["gross_limit"]:
        violations.append(f"GROSS weight: {axle_weights['gross']:,} lbs exceeds {TRAILER['gross_limit']:,} lb limit")

    return axle_weights, violations


# ---------------------------------------------------------------------------
# Trailer visualization widget
# ---------------------------------------------------------------------------

class TrailerView(QWidget):
    """Draws the trailer top-down with pallet slots labeled."""

    MARGIN = 30
    NOSE_LABEL_H = 24
    AXLE_LABEL_H = 18

    def __init__(self, parent=None):
        super().__init__(parent)
        self.plan: Optional[LoadPlan] = None
        self.setMinimumSize(300, 400)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def set_plan(self, plan: LoadPlan):
        self.plan = plan
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        self._draw(painter, self.rect())

    def _draw(self, painter: QPainter, rect):
        if self.plan is None or not self.plan.slots:
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "No load plan calculated yet.")
            return

        slots = self.plan.slots
        # Determine total rows
        max_row = max(s.row for s in slots)
        num_rows = max_row + 1

        m = self.MARGIN
        available_w = rect.width() - 2 * m
        available_h = rect.height() - 2 * m - self.NOSE_LABEL_H

        col_w = available_w / COLS
        row_h = available_h / num_rows

        trailer_rect = QRectF(m, m + self.NOSE_LABEL_H, available_w, available_h)

        # Trailer outline
        painter.setPen(QPen(QColor("#333"), 3))
        painter.setBrush(QBrush(QColor("#f5f5f0")))
        painter.drawRect(trailer_rect)

        # NOSE label
        font_label = QFont("Arial", 9, QFont.Weight.Bold)
        painter.setFont(font_label)
        painter.setPen(QColor("#333"))
        painter.drawText(QRectF(m, m, available_w, self.NOSE_LABEL_H),
                         Qt.AlignmentFlag.AlignCenter, "◀  NOSE (FRONT)")

        # Draw axle lines
        D = TRAILER["drive_axle_from_nose_in"]
        T = TRAILER["trailer_axle_from_nose_in"]
        trailer_len = TRAILER["interior_length_in"]

        def depth_to_y(depth_in):
            return trailer_rect.top() + (depth_in / trailer_len) * trailer_rect.height()

        for axle_depth, label, color in [
            (D, "Drive Axles", "#e67e00"),
            (T, "Trailer Axles", "#c0392b"),
        ]:
            y = depth_to_y(axle_depth)
            painter.setPen(QPen(QColor(color), 2, Qt.PenStyle.DashLine))
            painter.drawLine(QPointF(m, y), QPointF(m + available_w, y))
            painter.setPen(QColor(color))
            font_axle = QFont("Arial", 7)
            painter.setFont(font_axle)
            painter.drawText(QRectF(m + 2, y - 14, available_w - 4, 14),
                             Qt.AlignmentFlag.AlignRight, label)

        # Draw pallet slots
        colors = [
            QColor("#4a90d9"), QColor("#e67e22"), QColor("#27ae60"),
            QColor("#8e44ad"), QColor("#c0392b"), QColor("#16a085"),
            QColor("#d35400"), QColor("#2980b9"), QColor("#7f8c8d"),
            QColor("#f39c12"), QColor("#1abc9c"), QColor("#e74c3c"),
        ]

        font_slot = QFont("Arial", 8, QFont.Weight.Bold)
        font_small = QFont("Arial", 6)

        slot_map = {(s.row, s.col): s for s in slots}

        drawn_rows = set()
        for slot in slots:
            r, c = slot.row, slot.col
            x = trailer_rect.left() + c * col_w
            y = depth_to_y(slot.row_depth_in)

            if slot.piece:
                row_h_actual = (slot.piece.length_in / trailer_len) * trailer_rect.height()
            else:
                row_h_actual = row_h

            cell = QRectF(x + 1, y + 1, col_w - 2, row_h_actual - 2)

            color = colors[(slot.piece.piece_id - 1) % len(colors)] if slot.piece else QColor("#ccc")
            painter.setBrush(QBrush(color))
            painter.setPen(QPen(QColor("#222"), 1))
            painter.drawRect(cell)

            if slot.piece:
                pos_num = r * COLS + c + 1
                painter.setFont(font_slot)
                painter.setPen(Qt.GlobalColor.white)
                # Position number (large)
                painter.drawText(cell, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter,
                                 f"#{pos_num}")
                painter.setFont(font_small)
                desc = slot.piece.description[:10] if slot.piece.description else ""
                weight_str = f"{slot.piece.weight_lbs:,.0f} lbs"
                painter.drawText(cell, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignHCenter,
                                 f"[{slot.piece.piece_id}]\n{desc}\n{weight_str}")

        # Column headers (L / R)
        painter.setPen(QColor("#555"))
        font_hdr = QFont("Arial", 8)
        painter.setFont(font_hdr)
        for c, label in enumerate(["LEFT", "RIGHT"]):
            x = trailer_rect.left() + c * col_w
            painter.drawText(QRectF(x, trailer_rect.bottom() + 2, col_w, 16),
                             Qt.AlignmentFlag.AlignCenter, label)

        painter.end()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class FreightLoaderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Freight Loader — 53' Dry Van")
        self.setMinimumSize(1100, 720)
        self._next_id = 1
        self._plan: Optional[LoadPlan] = None
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setSpacing(8)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # ---- LEFT PANEL ----
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setSpacing(6)

        # Freight entry table
        grp_freight = QGroupBox("Freight Pieces")
        grp_layout = QVBoxLayout(grp_freight)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["#", "Description", "Length (in)", "Width (in)", "Height (in)", "Weight (lbs)"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        grp_layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("+ Add Piece")
        self.btn_add.clicked.connect(self._add_row)
        self.btn_del = QPushButton("Remove Selected")
        self.btn_del.clicked.connect(self._remove_row)
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_del)
        btn_row.addStretch()
        grp_layout.addLayout(btn_row)
        left_layout.addWidget(grp_freight)

        # Weight limits group
        grp_limits = QGroupBox("Axle Weight Limits (lbs)")
        limits_layout = QHBoxLayout(grp_limits)

        self._limit_fields = {}
        for key, label, default in [
            ("steer", "Steer", 12000),
            ("drive", "Drives", 34000),
            ("trailer", "Trailer", 34000),
            ("gross", "Gross", 80000),
        ]:
            vb = QVBoxLayout()
            lbl = QLabel(label)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            spin = QSpinBox()
            spin.setRange(1000, 200000)
            spin.setValue(default)
            spin.setSingleStep(500)
            spin.setGroupSeparatorShown(True)
            self._limit_fields[key] = spin
            vb.addWidget(lbl)
            vb.addWidget(spin)
            limits_layout.addLayout(vb)

        self.chk_override = QCheckBox("Allow override (ignore violations)")
        limits_layout.addWidget(self.chk_override)
        left_layout.addWidget(grp_limits)

        # Calculate button
        self.btn_calc = QPushButton("Calculate Load Plan")
        self.btn_calc.setFixedHeight(42)
        font_btn = self.btn_calc.font()
        font_btn.setPointSize(13)
        font_btn.setBold(True)
        self.btn_calc.setFont(font_btn)
        self.btn_calc.clicked.connect(self._calculate)
        left_layout.addWidget(self.btn_calc)

        # Axle weight summary
        grp_summary = QGroupBox("Axle Weight Summary")
        summary_layout = QVBoxLayout(grp_summary)
        self.summary_label = QLabel("—")
        self.summary_label.setWordWrap(True)
        self.summary_label.setTextFormat(Qt.TextFormat.RichText)
        summary_layout.addWidget(self.summary_label)
        left_layout.addWidget(grp_summary)

        # Violations / status
        grp_status = QGroupBox("Status / Violations")
        status_layout = QVBoxLayout(grp_status)
        self.status_label = QLabel("No plan calculated.")
        self.status_label.setWordWrap(True)
        self.status_label.setTextFormat(Qt.TextFormat.RichText)
        status_layout.addWidget(self.status_label)
        left_layout.addWidget(grp_status)

        # Print button
        self.btn_print = QPushButton("Print / Save PDF")
        self.btn_print.clicked.connect(self._print_plan)
        self.btn_print.setEnabled(False)
        left_layout.addWidget(self.btn_print)

        splitter.addWidget(left)

        # ---- RIGHT PANEL — trailer view ----
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)

        view_label = QLabel("Trailer Load View (Top-Down, Nose at Top)")
        view_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        view_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        right_layout.addWidget(view_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.trailer_view = TrailerView()
        self.trailer_view.setMinimumHeight(600)
        scroll.setWidget(self.trailer_view)
        right_layout.addWidget(scroll)

        splitter.addWidget(right)
        splitter.setSizes([480, 580])

        # Seed with a couple of example rows
        self._add_row(description="Pallet 1", length=48, width=48, height=60, weight=1800)
        self._add_row(description="Pallet 2", length=48, width=48, height=60, weight=2200)

    def _add_row(self, description="", length=48.0, width=48.0, height=48.0, weight=1000.0):
        row = self.table.rowCount()
        self.table.insertRow(row)

        id_item = QTableWidgetItem(str(self._next_id))
        id_item.setFlags(id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setItem(row, 0, id_item)

        self.table.setItem(row, 1, QTableWidgetItem(description))

        for col, val in enumerate([length, width, height, weight], start=2):
            item = QTableWidgetItem(str(val))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, col, item)

        self._next_id += 1

    def _remove_row(self):
        rows = sorted(set(i.row() for i in self.table.selectedItems()), reverse=True)
        for r in rows:
            self.table.removeRow(r)

    def _read_pieces(self) -> list[FreightPiece]:
        pieces = []
        for row in range(self.table.rowCount()):
            try:
                pid = int(self.table.item(row, 0).text())
                desc = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
                length = float(self.table.item(row, 2).text())
                width = float(self.table.item(row, 3).text())
                height = float(self.table.item(row, 4).text())
                weight = float(self.table.item(row, 5).text())
                pieces.append(FreightPiece(pid, desc, length, width, height, weight))
            except (ValueError, AttributeError) as e:
                QMessageBox.warning(self, "Input Error", f"Row {row+1} has invalid data: {e}")
                return []
        return pieces

    def _calculate(self):
        pieces = self._read_pieces()
        if not pieces:
            return

        # Apply user-edited limits to TRAILER dict
        TRAILER["steer_limit"] = self._limit_fields["steer"].value()
        TRAILER["drive_limit"] = self._limit_fields["drive"].value()
        TRAILER["trailer_limit"] = self._limit_fields["trailer"].value()
        TRAILER["gross_limit"] = self._limit_fields["gross"].value()

        override = self.chk_override.isChecked()
        plan = plan_load(pieces, override=override)

        if plan.violations and not override:
            msg = "Weight violations detected:\n\n" + "\n".join(plan.violations)
            msg += "\n\nCheck 'Allow override' to proceed anyway."
            QMessageBox.warning(self, "Weight Violations", msg)
            # Still show the plan but flag it
            self._plan = plan
            self.trailer_view.set_plan(plan)
            self._update_summary(plan)
            self.btn_print.setEnabled(True)
            return

        self._plan = plan
        self.trailer_view.set_plan(plan)
        self._update_summary(plan)
        self.btn_print.setEnabled(True)

    def _update_summary(self, plan: LoadPlan):
        aw = plan.axle_weights
        rows = []
        for key, label, limit_key in [
            ("steer",   "Steer",   "steer_limit"),
            ("drive",   "Drive",   "drive_limit"),
            ("trailer", "Trailer", "trailer_limit"),
            ("gross",   "Gross",   "gross_limit"),
        ]:
            val = aw.get(key, 0)
            limit = TRAILER[limit_key]
            pct = (val / limit * 100) if limit else 0
            color = "#c0392b" if val > limit else "#27ae60"
            rows.append(
                f"<tr>"
                f"<td><b>{label}:</b></td>"
                f"<td align='right'><span style='color:{color}'>{val:,} lbs</span></td>"
                f"<td align='right'>{pct:.0f}% of {limit:,}</td>"
                f"</tr>"
            )
        self.summary_label.setText("<table>" + "".join(rows) + "</table>")

        if plan.violations:
            viol_html = "<br>".join(
                f"<span style='color:red'>⚠ {v}</span>" for v in plan.violations
            )
            status = viol_html
            if plan.override:
                status += "<br><i>Override active — proceeding despite violations.</i>"
        else:
            status = "<span style='color:green'>✓ All axle weights within limits.</span>"
        self.status_label.setText(status)

    def _print_plan(self):
        if self._plan is None:
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Load Plan PDF", "load_plan.pdf", "PDF Files (*.pdf)"
        )
        if not path:
            return

        writer = QPdfWriter(path)
        writer.setPageSize(QPageSize(QPageSize.PageSizeId.Letter))
        writer.setPageOrientation(QPageLayout.Orientation.Portrait)
        writer.setResolution(150)

        painter = QPainter(writer)

        page_rect = painter.viewport()
        margin = 80
        usable = page_rect.adjusted(margin, margin, -margin, -margin)

        # Title
        font_title = QFont("Arial", 18, QFont.Weight.Bold)
        painter.setFont(font_title)
        painter.drawText(usable, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter,
                         "Freight Load Plan — 53' Dry Van")

        # Axle weights text
        aw = self._plan.axle_weights
        font_body = QFont("Arial", 11)
        painter.setFont(font_body)
        lines = [
            f"Steer: {aw.get('steer', 0):,} lbs   Drive: {aw.get('drive', 0):,} lbs   "
            f"Trailer: {aw.get('trailer', 0):,} lbs   Gross: {aw.get('gross', 0):,} lbs"
        ]
        if self._plan.violations:
            lines += ["VIOLATIONS: " + "; ".join(self._plan.violations)]
        if self._plan.override:
            lines += ["** Override active **"]

        y_offset = 120
        for line in lines:
            painter.drawText(
                usable.adjusted(0, y_offset, 0, 0),
                Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft,
                line
            )
            y_offset += 60

        # Trailer drawing in remaining space
        draw_rect = usable.adjusted(0, y_offset + 20, 0, 0)
        qrectf = QRectF(draw_rect)
        self.trailer_view._draw(painter, qrectf)

        painter.end()

        QMessageBox.information(self, "Saved", f"Load plan saved to:\n{path}")


# ---------------------------------------------------------------------------
# Manifest table dialog (position list)
# ---------------------------------------------------------------------------

class ManifestDialog(QDialog):
    def __init__(self, plan: LoadPlan, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Load Manifest")
        self.setMinimumSize(480, 400)
        layout = QVBoxLayout(self)

        lbl = QLabel("<b>Pallet Position Manifest</b>")
        layout.addWidget(lbl)

        table = QTableWidget(0, 5)
        table.setHorizontalHeaderLabels(["Position", "Side", "Freight #", "Description", "Weight (lbs)"])
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        for slot in sorted(plan.slots, key=lambda s: (s.row, s.col)):
            r = table.rowCount()
            table.insertRow(r)
            pos = slot.row * COLS + slot.col + 1
            side = "Left" if slot.col == 0 else "Right"
            table.setItem(r, 0, QTableWidgetItem(str(pos)))
            table.setItem(r, 1, QTableWidgetItem(side))
            if slot.piece:
                table.setItem(r, 2, QTableWidgetItem(str(slot.piece.piece_id)))
                table.setItem(r, 3, QTableWidgetItem(slot.piece.description))
                table.setItem(r, 4, QTableWidgetItem(f"{slot.piece.weight_lbs:,.0f}"))
            else:
                for c in [2, 3, 4]:
                    table.setItem(r, c, QTableWidgetItem("(empty)"))

        layout.addWidget(table)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = FreightLoaderApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
