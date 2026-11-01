# Freight Loader — Project Notes

_Last updated: 2026-10-26_

## Where we are

- **Version:** 2.1.0 — committed on branch **`v2.1.0`**, merged to `main` and pushed
  to GitHub (`jaymc97/freight-loader`). New axle-weight model with fifth wheel / tandem
  positions, and the operator stack-block row order.
- **v2.0.0 is kept** to revert if needed:
  - plain copy of the files in `Previous Versions\v2.0.0\` (local only — git-ignored)
  - git tag `release-2.0.0` and branches `v2.0.0` / `backup/v2.0.0` (commit `629ae41`)
  - To go back: `git checkout release-2.0.0`, or copy
    `Previous Versions\v2.0.0\freight_loader.py` over.

## Run it

```
cd "C:\Users\jaymc\Downloads\Local App Ideas\Freight-Loader"
python freight_loader.py
```
Requires `PyQt6` and `openpyxl` (`pip install -r requirements.txt`). PDF preview uses
PyQt6's QtPdf module (included in the installed PyQt6).

## Folders

| Folder / file | Contents |
|---|---|
| `Load Plans\` | Saved plans (`.json`) and exported PDFs. Ignored by git. |
| `Saved Load Templates\` | Filled-in Excel packing lists (was `Packing Lists`; the app renames it on startup if needed). Ignored by git. |
| `Freight_Loader_Import_Template.xlsx` | Blank import template (regenerate with `python make_template.py`) |
| `freight_loader.py` | The whole app |
| `data\` | Trailer types, state size/weight limits, US state boundaries (from the Cargo Load Planner app) |
| `generate_traffix_plan.py` | Old helper script that builds one plan and its PDFs |

Save dialogs start in the last folder used (Windows setting `FreightLoader/last_save_dir`),
defaulting to `Load Plans`.

## What the app does (v2.1.0)

- **Planner** (`plan_load`): pallets narrower than half the trailer pair two across,
  matched by length; wider crates take a full row. Stacks only until the load fits.
  Anything that won't fit is reported (last-delivered stop's pieces are left off first)
  and never drawn outside the trailer.
- **Row order (v2.1.0, operator rule):** within each stop zone the stacked (double) rows
  form **one block in the middle**; with an odd number of single rows the extra single
  goes behind the block, i.e. the block shifts toward the nose (12 rows, 3 doubles →
  4 singles, 3 doubles, 5 singles — how operators load the Abezzol plan). Rows then only
  trade places single↔single / double↔double to even the axles. If that's not legal even
  with the tandem / fifth wheel slid, the block moves the fewest rows needed; only if no
  block position is legal may the stacks split up (`_balanced_order`).
- **Balance:** row order keeps axles legal with drive and trailer axles even. Left/right is set by `_balance_sides`: wide wall-hugging crates
  **alternate walls row by row** (start right on ties; tail crate may be centred);
  pairs pick sides to balance. It also checks balance after each delivery.
- **Multi-stop** (stop IDs 1–10): zones nose→doors by delivery order; pieces only pair or
  stack within a stop (except one shared boundary row); axle checks after each delivery.
- **Multi-truck:** split by stop along the route. When a stop won't fit, the user chooses
  "move whole stop to a new truck" or "load what fits, ship the rest separately".
  Truck bar: selector, Re-split, Move Stop…, Ship Separately list.
- **Stops & route:** Edit Stops… takes an address (OpenStreetMap lookup, online) or GPS
  coordinates / Google Maps link (offline) for the origin and each stop. Nearest-first +
  2-opt routing; straight-line by default, road miles (OSRM, online) optional.
- **Views:** to-scale Top and Side views (zoom: −/Fit/+, Ctrl+wheel, right-drag pan),
  rotatable 3D view with presets.
- **Top View editing:** drag a crate to shift it (stops on contact, snaps within 2");
  Ctrl+drag swaps two side-by-side crates; Reset Moves. Weights recalculate from actual
  positions. Moves are saved in the plan.
- **Securement** (3D tab): load bars, straps (up to 2 per face), airbags. Airbags follow
  the gap; straps are refused on faces touching another crate; hidden items can be
  cleared with "Remove Hidden (n)".
- **Output:** Export buttons preview the PDF before saving. The floor plan covers every
  truck, with a side elevation and securement checklist, an optional 3D page and a
  ship-separately page. Crate labels show STOP and Truck.
- **Excel:** "Open Excel Template…" creates a new list in `Saved Load Templates`. Column L
  of the template is Stop #.
- **Axle weights (v2.1.0):** two-beam model (`_axles_for`). Trailer rests on the kingpin
  and its tandem; carries freight + empty body (at its balance point); the slider's weight
  sits on the tandem. Kingpin load splits drive/steer by where the fifth wheel sits ahead
  of the drive tandem. Empty rig defaults 34,000 lb (tractor 10,800 steer + 8,700 drive,
  trailer 14,500) → legal payload 46,000 at 80,000 gross (`max_payload()`).
- **Fifth wheel / tandem:** placed automatically per truck (`auto_rig`): fifth wheel at its
  usual 12", tandem as far back as the kingpin-to-rear-axle limit allows (41' default
  since 2026-10-31 → 462" tandem centre; was 40' → 450"); slides only when drive or trailer is over 90% or an axle (incl.
  after each delivery) would be overweight. The user can move them in the **Axles** row
  (spin boxes, snap to 2" notches / 6" holes) or by **dragging the wheels in the Side
  View**; **Auto** hands control back. Manual positions are kept per truck and saved.
  Rear axle past the limit is flagged as a violation. **Truck & Trailer…** edits empty
  weights, wheelbase, slide ranges and the limit (saved as app default and in each plan).
- Side View shows kingpin ▲, drive and trailer wheels with their weights, steer weight at
  the nose; Top/3D draw the drive and trailer axle lines at their real positions. PDF
  footer gives the positions to set and the gross.

### Added 2026-10-31 (not yet committed) — ported from the Cargo Load Planner app

Source: the Flutter app `com.luminarylogistics.cargo_load_planner` (on the phone), project at
`C:\Users\jaymc\OneDrive\Desktop\Inventory Manager OneDriveX\Cargo Load Planner`.
Data files copied to `data\` (`.gitignore` has `!data/*.json`).

- **Route & Permits…** (beside Edit Stops): each truck's road route (OSRM, origin → its
  stops in delivery order, online) → miles, drive hours, driving day each stop is reached
  (11 h/day, FMCSA hours of service), **states crossed** (route line vs
  `data\us_states_boundaries.geojson`; a state counts after 0.5 mi so coarse borders don't
  add states), and each state's **legal / permit limits** (`data\state_regulations.json`)
  vs the truck's overall width, height, trailer length and gross. Results are cached per
  route in the plan file (`shipment_plan.routes`), shown in the Axle Weight Summary and
  Status, and added as a "Route & Permits" page at the end of the floor plan PDF.
  **The state data is AI-researched and not verified** — every record says so; e.g. it
  lists CT's legal length as 48' (KPRA state), so a 53' van gets flagged there.
- **Trailer types** (`data\trailer_types.json`): 53' dry van, 48'/53' flatbed, 48'
  step-deck, Conestoga, extendable step-deck, double-drop/lowboy, 3-axle RGN. Each truck
  has its own type (Trailer row above Axles; **Auto** hands it back). Freight that fits a
  van is split into vans exactly as before; pieces that don't go to the first open-deck
  type that carries them legally, else the least oversize one (`split_shipment`,
  `best_trailer_for`). An oversize/overweight piece rides **alone** (permits are for one
  indivisible piece); the planner never stacks or combines into oversize. Open decks take
  overhanging pieces; oversize width/height/length and overweight are flagged in Status,
  on the PDF and per state in Route & Permits. On open decks the load slides back along
  the deck when the axles are overloaded at the nose (`_slide_load_back`).
  `set_trailer_type()` / `with trailer_type(id)` switch the global TRAILER + trailer half
  of RIG; each type keeps its own running gear (Truck & Trailer… edits the current
  truck's type; saved in settings `rig_trailers` and in the plan file).
- **Quick Add…** (next to + Add Row): N identical pieces from one size/weight entry;
  "This is all the freight for the load" ticks **All freight entered** and calculates.
- **All freight entered — spread the load** (DOT limits box, saved per plan as
  `shipment_plan.spread`): if packing two-across from the nose overloads an axle,
  `_spread_rows` turns side-by-side rows (heaviest first) into single centred rows until
  every axle is legal, never adding stacks. Off = freight stays at the nose (more coming);
  a status note suggests ticking it when axles are over and floor is free.
  Test: `Load Plans\15 identical 2026-10-31.json` (15 × 48" cubes, 2,900 lb) — nose-packed
  steer 14,064 / drive 36,712; spread 11,906 / 33,028 / 32,567 (5 single rows, 40').
- **Top View moves:** dragging snaps to the trailer's centre line within 4". When the load
  changes, only moves that no longer fit are put back (not all of them). Pos. → Center on a
  moved crate centres it where it is (refused, with a note, if a crate is beside it).

## Next time — to look at

1. **Try the axle controls in the real app** (only tested headless so far), and enter real
   scale tickets / measurements in Truck & Trailer… if the defaults don't match the fleet.
2. **Excel import checks** (from the 2026-10-26 review, not done yet): blank L/W/H/weight
   import as 0 with no warning; duplicate crate #s aren't caught; one bad cell (e.g. "12 pcs")
   aborts the whole import; Browse offers `.xls` that openpyxl can't read; crates unticked at
   import can't be re-included later.
3. **Move Stop can select the wrong truck** after an emptied truck is removed (`_move_pieces`).

4. **Stop column in the Pop Out.** It was added this session: "Stop" between PCS and
   Notes, showing e.g. "S2 — Coldwater, MI" in the stop's colour. Check it in the real
   app and adjust if it's not what's wanted (editable? sortable? filter by stop?).
5. **Ability to unstack pieces.** Let the user say "don't stack this one" or break an
   existing stack and have the planner re-plan around it. Ideas to discuss:
   - per-piece "Never stack on top / never stack under" options (today: the "Stack?"
     checkbox only means "can carry something on top"; a piece marked N can still ride on top)
   - right-click a stacked crate in the Top/3D view → "Unstack", then re-plan (it needs
     floor space, so it may push something to another truck or ship-separately)
   - pinning user decisions so a re-plan doesn't undo them
6. **Heavy-haul trailers (left out on purpose, 2026-10-31).** The Cargo app's "Multi-Axle
   Heavy Haul RGN (9+ axle)" (150,000 lb) wasn't brought over: Freight Loader's axle model
   is tractor + one trailer axle group. Pieces too big/heavy for every type are sent to
   Ship Separately as "needs a heavy-haul carrier". Supporting it means multiple axle
   groups (jeeps/boosters/stingers), per-group limits and superload permits.
7. **Trailer-type follow-ups:** step-decks are modelled as one deck at the main-deck
   height (upper deck ignored); lowboy/RGN as the well only, kingpin at the well front;
   the extendable step-deck's tandem doesn't move with the extension; overhang past the
   deck sides isn't drawn (it's reported); 3D/Side views still draw a box at the legal
   cargo height on open decks; trailer empty weights / axle positions are typical guesses.
8. **Truck routing:** OSRM routes as a car (drive times are car speeds). The Cargo app
   used OpenRouteService's HGV profile (needs a free API key) — option for later.

## Assumptions worth revisiting

- 53' dry van: 636" × 98" × 110" interior. Freight positions are measured from the
  interior nose; kingpin 36" behind it.
- Truck & trailer defaults (`RIG_DEFAULTS`) are typical figures, not measured — the empty
  body balance point (270" from nose) and slider weight (3,000 lb) set the empty split
  (≈5,600 kingpin / 8,900 tandem).
- Row order is optimised with the fifth wheel / tandem at their default spots, then the
  axles are slid; the planner doesn't re-order rows when you move the axles by hand.
- Fifth wheel stays put between deliveries (after-delivery checks use the same positions).
- Kingpin-to-rear-axle default is 41' (492") since 2026-10-31, except a truck whose route
  goes through **California** is held to 40' (480") automatically (`kpra_limit()`,
  `_route_kpra_for`): its looked-up route (Route & Permits) decides; before a lookup, the
  origin or a stop in California does. Looking up routes re-plans trucks whose limit
  changes. The Axles row shows "(CA 40' max)". Plans saved earlier carry 40' in their own
  truck specs and keep it when opened. In random 3-stop tests near full payload (43–46k),
  several loads were 300–2,500 lb over on the trailer tandem at 40' and about half of those
  became legal at 41'. When axles are over only because of this limit, Status says which
  tandem / fifth wheel setting past it would be legal (`kpra_hint`).
- Stack block rule assumes "doubles toward the centre" applies per stop zone, also to
  side-by-side pallet rows where either pallet is stacked.
- "Stackable = N" means nothing may go on it; it can still go on top of another piece.
- The planner doesn't re-balance around crates dragged by hand.
- Row order follows axle balance within each stop zone, not list order.

## Test files

- `Saved Load Templates\test_load_random.xlsx`: 36 pcs, 44,010 lb, one stop (fits one truck)
- `Saved Load Templates\test_load_multistop.xlsx`: same load over 4 stops (needs 2 trucks)
- `Load Plans\Abezzol_JM-8675309_2026-10-24.json` / `Corrected Abezzol…json`: 15
  identical wide crates; the corrected version shows the required alternating pattern
- `Load Plans\Multi-Stop Test 2026-10-24.json`: 2-truck plan used for the balance review
- Corrected Abezzol / Plan1: planner now gives 4 singles, 3 doubles, 5 singles (drive
  26,671 / trailer 29,290); at 3,050 lb per crate the centred block isn't legal, so it moves
  to the nose and the tandem slides to 426" (drive 33,874 / trailer 33,932).
- v2.1.0 results on these (auto positions, before the stack-block rule): Solar6 33,750 lb → 67,750 gross, drive 27,903 /
  trailer 27,990; random list 44,010 lb on one truck → 78,010 gross, drive 33,160 / trailer
  32,938 (fifth wheel 10"). Same load packed heaviest-first at the nose: default positions
  put steer 12,214 and drive 34,381 over — auto slides tandem to 426" and fixes it.
