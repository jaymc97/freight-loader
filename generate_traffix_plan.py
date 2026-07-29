"""
Generates load plan PDF + crate labels PDF for the Traffix shipment (Crates 1-25).
Run: python3 generate_traffix_plan.py
Outputs: Traffix_Load_Plan.pdf, Traffix_Crate_Labels.pdf
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from PyQt6.QtWidgets import QApplication
from freight_loader import (
    FreightPiece, ShipmentInfo, plan_load,
    export_load_plan_pdf, export_crate_labels_pdf
)

app = QApplication(sys.argv)

CRATES = [
    FreightPiece(1,  "Crate 1",  "NE41P-0064-B", 42, 51, 18, 1000, 500),
    FreightPiece(2,  "Crate 2",  "NE21P-0079-C", 42, 42, 43, 1650, 270),
    FreightPiece(3,  "Crate 3",  "NE21P-0080-B", 43, 51, 17,  700, 225),
    FreightPiece(4,  "Crate 4",  "NE21P-0083-B", 48, 43, 37, 1765,  18),
    FreightPiece(5,  "Crate 5",  "NE21P-0083-B", 48, 43, 37, 1765,  18),
    FreightPiece(6,  "Crate 6",  "NE21P-0083-B", 48, 43, 37, 1765,  18),
    FreightPiece(7,  "Crate 7",  "NE21P-0083-B", 48, 43, 37, 1765,  18),
    FreightPiece(8,  "Crate 8",  "NE21P-0083-B", 48, 48, 43,  785,   8),
    FreightPiece(9,  "Crate 9",  "NE21P-0089-A", 42, 42, 43, 2000, 284,
                 notes="Approx. weight"),
    FreightPiece(10, "Crate 10", "NE22P-0071-D", 44, 44, 55, 1700, 880,
                 notes="Also contains 120+240 pcs NE23P-0068-B"),
    FreightPiece(11, "Crate 11", "NE22P-0071-D", 44, 44, 68, 2000, 1620),
    FreightPiece(12, "Crate 12", "NE41P-0064-B", 42, 51, 18,  900, 436,
                 notes="Remaining pcs per Traffix instruction"),
    FreightPiece(13, "Crate 13", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(14, "Crate 14", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(15, "Crate 15", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(16, "Crate 16", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(17, "Crate 17", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(18, "Crate 18", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(19, "Crate 19", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(20, "Crate 20", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(21, "Crate 21", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(22, "Crate 22", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(23, "Crate 23", "NE22P-0071-D", 44, 44, 71, 2000, 1620),
    FreightPiece(24, "Crate 24", "NE22P-0071-D", 48, 44, 71, 2000, 1620),
    FreightPiece(25, "Crate 25", "NE22P-0071-D", 48, 44, 71, 2000, 1620),
]

INFO = ShipmentInfo(
    customer="Traffix",
    shipment_name="Traffix Shipment",
    date="2026-07-29",
    trailer_num="",
    bol_num="",
)

EXCLUDED = [26, 27]

plan = plan_load(CRATES, excluded_ids=EXCLUDED)

out_dir = os.path.dirname(__file__)
plan_path  = os.path.join(out_dir, "Traffix_Load_Plan.pdf")
label_path = os.path.join(out_dir, "Traffix_Crate_Labels.pdf")

export_load_plan_pdf(plan, INFO, plan_path)
export_crate_labels_pdf(plan, INFO, label_path)

aw = plan.axle_weights
print("=== Traffix Load Plan ===")
print(f"Total cargo: {sum(c.weight_lbs for c in CRATES):,.0f} lbs across {len(CRATES)} crates")
print(f"Excluded: Crates {EXCLUDED}")
print(f"Trailer length used: {plan.total_length_used_in:.0f}\" of {636}\" "
      f"({plan.total_length_used_in/12:.1f} ft of 53 ft)")
print()
print("Load Order (nose → tail):")
slot_map = {(s.row, s.col): s for s in plan.slots}
rows = sorted(set(s.row for s in plan.slots))
for r in rows:
    a = slot_map.get((r, 0))
    b = slot_map.get((r, 1))
    a_str = f"{a.piece.crate_label} ({a.piece.weight_lbs:,.0f} lb)" if a and a.piece else "VOID"
    b_str = f"{b.piece.crate_label} ({b.piece.weight_lbs:,.0f} lb)" if b and b.piece else "VOID"
    print(f"  R{r+1:>2}  A: {a_str:<28}  B: {b_str}")

print()
print("Estimated Axle Weights:")
print(f"  Steer:          {aw.get('steer', 0):>8,} lbs  (limit 12,000)")
print(f"  Drive:          {aw.get('drive', 0):>8,} lbs  (limit 34,000)")
print(f"  Trailer Tandem: {aw.get('trailer', 0):>8,} lbs  (limit 34,000)")
print(f"  Gross (est.):   {aw.get('gross', 0):>8,} lbs  (limit 80,000)")

if plan.violations:
    print("\n⚠  VIOLATIONS:")
    for v in plan.violations:
        print(f"   {v}")
else:
    print("\n✓ All axle weights within legal limits.")

print(f"\nPDFs written:")
print(f"  {plan_path}")
print(f"  {label_path}")
