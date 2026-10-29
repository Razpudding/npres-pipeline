# RES West-Overijssel — Monitor → GeoPackage sync

Tooling to update the NP RES _projecten-geopackage_ (layers `npres_zov_v4` — Zon op Veld, and `npres_wol_v4` — Wind op Land) from the regional monitor Excel of RES West-Overijssel.

The RES region tracks solar-field and onshore-wind projects in an Excel monitor with its own column names, value conventions and phase model. NP RES requires the same data as a GeoPackage that follows the national _invulformat_ data model (fixed column names, field lengths and choice lists). This repo bridges the two: it documents the column mapping between both models and provides a script that applies monitor updates to the GeoPackage in a controlled, reviewable way.

## Contents

| File                                    | Purpose                                                                                                                                                                                                                                                                                                               |
| --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `update_res_gpkg.py`                    | Two-phase update script (plan → apply)                                                                                                                                                                                                                                                                                || `NP_RES_data_portaal_instructies_*.pdf` | NP RES documentation of the GeoPackage data model (attribute descriptions per layer)                                                                                                                                                                                                                                  |

## Usage

Put the source GeoPackage and monitor Excel in the `input/` folder; all output (the plan CSVs and the updated GeoPackage) is written to `output/`. Filenames passed to `--gpkg` / `--excel` are resolved against `input/` (a full path also works), and the input files are never modified. Both folders are git-ignored except for a `.gitkeep`, so data is never committed.

```bash
# dry run: compute and report all mutations, change nothing
python update_res_gpkg.py --gpkg projecten.gpkg --excel monitor.xlsx

# execute (after reviewing the plan); asks for confirmation, then writes an
# updated *copy* to output/ — the input gpkg is left untouched
python update_res_gpkg.py --gpkg projecten.gpkg --excel monitor.xlsx --apply

# the monitor sheet is auto-detected (first sheet starting with "Monitor");
# override with --sheet "Monitor Q2 2026", or the folders with --input-dir/--output-dir
```

Requirements: Python 3.9+, `pandas`, `openpyxl`. **No GDAL/GeoPandas needed** — a GeoPackage is a SQLite database, and the script talks to it directly via the stdlib `sqlite3` module.

Every plan run writes three CSVs to the `output/` folder for review:

- `plan_mutations.csv` — every intended change: layer, fid, project id, column, old value, new value
- `plan_warnings.csv` — everything that needs a human decision (unparseable values, refused overwrites, uncertain mappings)
- `plan_coverage.csv` — per project: matched / new in Excel / only in gpkg / stopped

## How the two datasets relate

- **Join key:** Excel column `Nr.` (integer) ↔ gpkg `zv_prj_id` / `wl_prj_id` as `"ES05_" + Nr`. Leading zeros in Excel are display formatting only, so `Nr` is cast to `int` before concatenating. The join is per layer: Excel `Type` routes rows to `npres_zov_v4` (`Zon-op-veld`) or `npres_wol_v4` (`Wind`); `Zon-op-dak` rows are out of scope.
- **Snapshot difference:** the gpkg represents an earlier reporting round (Q4 2025); the monitor is current (Q2 2026). Most value differences are therefore legitimate progress: phases advanced, Slaagkans changed, production figures updated.
- **The Excel is leading** for all mapped attribute columns. The gpkg remains leading for everything the monitor doesn't track (geometry, trede, opslag, transportcapaciteit, …) and for a small set of "special case" rows that intentionally exist only in the gpkg — those are reported but never touched.

## Project-ID history (`npres_wol_v4`)

**Project ids are per layer.** `ES05_116` exists independently in `npres_zov_v4` (Zandwinplas Hasselt, GM1896) and in `npres_wol_v4` — they are unrelated projects that happen to share a number. Never join across layers on the id alone.

Ids in the wind layer have been renumbered and pruned by hand, so a bare id is not stable over time. The mapping below is the record of that; without it, the same project looks like two different ones.

| gpkg id (old) | now | Project | What happened |
| ------------- | --- | ------- | ------------- |
| `ES05_116` | **`ES05_123`** | Windpark Groen (Steenwijkerland) | Renumbered between 2026-07-07 and 2026-07-13. Its geometry was the entire 321 km² gemeente polygon; replaced 2026-07-20 with a 10 × 10 m placeholder (`..._prebackup_fixgeom116_...`). `wl_gemcode` was the wrong `GM1896` — corrected to `GM1708`. |
| `ES05_127` | *(deleted)* | Haarst-De Ruit | `ONLY_IN_GPKG`, no monitor row. Deleted 2026-07-20 (`..._prebackup_delete127_133_...`). Sits 2.3 km from ES05_105, so it reads like a superseded Genne-Haerst delineation — but it is **not** the same project as ES05_105. |
| `ES05_133` | *(deleted)* | Deventer | `ONLY_IN_GPKG`, no monitor row. Deleted 2026-07-20, same step. |
| `ES05_110` | *(deleted)* | Dalfserveld-west (deel Wind) | `Project gestopt` in the Q2 monitor → `STOPPED_DELETED`. |
| — | `ES05_94` | Broeklanden (Hardenberg) | `NEW_IN_EXCEL` in Q2 2026. No X/Y in the monitor → inserted with NULL geometry, still to be digitised. |
| — | `ES05_119` | Kloosterhaar (Hardenberg) | `NEW_IN_EXCEL` in Q2 2026. Same: NULL geometry, still to be digitised. |

**Placeholder rows carry the gemeente name, not the project name.** Where the monitor has a real name, the gpkg may still hold a bare municipality: `Kampen` = Windpark N50 (ES05_098), `Staphorst` = Windpark Rouveen-Oost (ES05_118), `Steenwijkerland` = Windpark Groen (ES05_123). Matching on name will therefore fail — always match on `"ES05_" + int(Nr)`.

**The `data_aanlevering_WO_*.gpkg` branch.** The hand-maintained aanlevering GeoPackage under `NPRES Dataportaal/` descends from this lineage but branched *before* the 2026-07-20 corrections, so it carried the pre-renumber ids for months. It is not redundant: it holds the per-turbine digitising (one row per turbine, ~47 rows) that the pipeline output never had (one placeholder row per project). Reconciled against the Q2 monitor on 2026-09-24 — the wind-layer project set now matches both the monitor and `output/projecten_geopackage_2026_0_updated.gpkg` exactly. When diagnosing discrepancies, diff against that file first; the output filenames double as a log of the manual steps taken.

**Wind-layer attributes synced to the Q2 monitor on 2026-10-19** (the zov layer had been done on 2026-09-24). Changed: `ES05_95` 60 → 42 MW and 10 → 7 turbines (`vol_r` 3333, `pr`/`pr_plam` recomputed); `ES05_114` `reagr_r` 0.5 → 0.1 (`pr`/`pr_plam` recomputed); `ES05_32`/`ES05_33` now in production — `fase` → `huidig`, `reagr_r` 1.0, `pr` 9.9, `pr_plam` 0, `dt_vv` 2023-04-01, `dt_rea` 2025-12-01; `wl_ondstat` filled from the monitor (left empty where the monitor says `nb`/`nnb` or is blank: ES05_95, ES05_114, ES05_123). Net effect: `pr` −37.1 GWh, `pr_plam` −23.9 GWh. **Deliberately left as-is:** ~1 % `vol_r`/`pr` rounding differences against the monitor (ES05_115/117/118/098/32/33), the gemeente-name placeholders above, and ES05_62's turbine count and heights (blank in the monitor — blanks never clear gpkg values).

## Column mapping (summary)

The mapping is encoded directly in the script — `APPLY_COLUMNS` plus the `transform_row` logic — which is the source of truth. (It was originally derived from a column-mapping workbook that is **not** part of this repo, as it contains sensitive data.) The short version:

**Direct copies:** `Projectnaam→*_naam`, `Vermogen→*_verm`, `Slaagkans→*_reagr_r`, `Prognose_GWh→*_pr` (its Excel formula, Slaagkans × Productie, is exactly the NP RES definition of `pr`), `Onderstation→*_ondsta(t)`, `Aantal turbines→wl_aant_t`, `Ashoogte→wl_masth`, `Tiphoogte→wl_tiph`, `Oppervlak zonnepark→zv_opp_pan`.

**Value maps / transformations:**

- `Fase`: the monitor's labels are mapped onto the NP RES choice list, absorbing the monitor's spelling quirks (`Productie→huidig`, `Vergunningsaanvraag→vergunningaanvraag`, `Subsidiebeschiking en bouw` [sic, one *k*] `→subsidiebeschikking en bouw`).
- `Gemeente` (name) → CBS `GM####` code via a hardcoded lookup of the eleven West-Overijssel municipalities, verified empirically against the gpkg.
- `F3/F4/F6` → `*_dt_va` / `*_dt_vv` / `*_dt_rea`. The F-columns are "date status _n_ started" per the monitor's own column spec; F3/F4/F6 correspond to the vergunningaanvraag/vergunningverlening/productie phases. Month-year precision, stored as ISO dates.
- `Lokaal eigendom`: a fraction in Excel (with pollution like `"100%"`, `"25%"`, `" -"`), a 0–100 percentage in the gpkg. Parsed defensively, ×100, and additionally derived into the `*_le` Ja/Nee flag.
- `Vorm lokaal eigendom`: monitor digits → NP RES letter codes. **The mapping is not positional**: monitor 1 = financiële participatie zonder eigendom → `A (ze)`, 2 = omgevingsfonds → `C (fon)`, 3 = omwonendenregeling → `B (omw)`, and 4 = procesparticipatie has no clean NP RES equivalent — it is mapped to `D (and)` and every occurrence emits a _verify_ warning. Combination strings (`"2,3,4"`, `"1 en 4"`) are parsed and rendered as the corresponding NP RES combination label (`BCD (omw+fon+and)`).
- `rescode` is the constant `ES05`.
- `SDE→*_subs_id`: only copied when the monitor cell actually contains an SDE code (e.g. `SDE1913304`, including multi-code `"… en …"` strings). Non-reference values (`Ja`, `Nee`, `Per turbine(set)`, `deels`, `0`) are ignored and leave the existing gpkg value untouched.

**Derived (no Excel source column):** `*_vol_r = round(Productie / Vermogen × 1000)` and `*_pr_plam = Productie × (1 − Slaagkans)`, both validated against the existing gpkg values before adopting the formulas.

The official NP RES definition is `*_pr_plam = *_verm × *_vol_r / 1000 × (1 − *_reagr_r)` — i.e. from the *stored* `vol_r`, which is rounded to whole hours. The script's `Productie`-based formula therefore differs from it by a rounding margin (well under 1 %); accepted as-is. Hand edits in the aanlevering gpkg use the official formula, so a pipeline run can shift those values by that margin.

**Deliberately unmapped:** `zv_trede` (the monitor's "Type zonnepark" values were tested against it — no correlation; they encode the provincial zonneladder, not the national voorkeursvolgorde), `*_bev_gez` (the monitor's "Omgevingsvergunning" is a Ja/Nee _granted_ flag, not the bestuurslaag), `zv_jr_vrea` (source column is nearly empty and contains junk like Excel-serial dates and `"22/23"` — disabled by default), and everything the monitor simply doesn't track (`tra_cap`, `opslag`, `rol`, `bgt_id`, `brt_id`, `verm_om`, `le_part`).

### `Vollasturen` column in the monitor Excel (added 2026-10-20)

In `RES West monitor_data_26_Q4.xlsx`, sheet `Data Veld+Wind (Q2'26)`, a `Vollasturen` column (V) was inserted between `Vermogen` (U) and `Productie` (W), so that `Productie [GWh] = Vermogen [MW] × Vollasturen [h] / 1000`. `Productie` itself was **not** changed. The column was inserted in Excel (so all references in `Monitor`, pivots and slicers shifted correctly); values were then written cell-by-cell. Every written cell, and the header, has a **red background**. Projects with status `0 Project gestopt` were left empty.

How each value was derived:

| Layer | Rule | Count (active projects) |
| ----- | ---- | ----------------------- |
| Wind | Formula `=IF(AND(U>0,W>0),ROUND(W/U*1000,0),"")` — same definition as gpkg `wl_vol_r` | 22 |
| Zon, group A | `Productie` is a formula `=U×0.9` or `=U×0.95` → the factor × 1000 (**900** / **950**), written as a fixed value | 10 × 900, 2 × 950 |
| Zon, group B | `Productie` is a fixed value or an area-based formula (`=ha × 0.8182`, with `Vermogen` back-calculated as `Productie / 0.9`), and `Productie / Vermogen × 1000` is exactly 900 or 950 → that value | 14 × 900, 16 × 950 |
| Zon, group C | `Productie / Vermogen × 1000` is not a standard value — most have a rounded, presumably real `Productie` estimate → the **rounded ratio**, written as-is | 12 (see below) |

Group C (non-standard, kept deliberately): ID 41 Zonneweide Lierderholthuis 833, 60 Oldemarkt 833, 7 Waterberging Broekland 886, 122 Eeser Campus 899, 116 Zandwinplas Hasselt 938, 13 Zonnepark Lus afrit 23 A1 952, 12 Locatie Waterschap 953, 5 Zandwinplas Luttenberg 954, 104 Energielandgoed Raalte 957, 22 Zonnepark Hardenberg 964, 37 Zonnepark en dak Vollenhove waterschap 967, 54 Zonnepark Weekhorst 972.

Solar values are **fixed numbers**, not formulas: if `Vermogen` or `Productie` changes for a solar project, update its `Vollasturen` by hand. The script still derives `*_vol_r` from `Productie / Vermogen`; it could read this column directly in future (it matches on header names, so the inserted column does not break it).

Other red cells in the same sheet (same date, copied from the aanlevering gpkg): `X_coordinaat`/`Y_coordinaat` for 25 projects (12 filled blanks — 5 of them placeholder locations, see the digitising queue — and 13 corrected by 100–711 m, using the centre of the gpkg geometry), and `Aantal turbines`/`Ashoogte`/`Tiphoogte` for ES05_62 Tolhuislanden (4 / 85 / 125.5 — identical to ES05_2 Westenwind, unverified).

## Design decisions in the script

**Two-phase execution.** The plan phase computes every mutation against in-memory copies of both sources and reports them (counts per column, warnings, coverage) before anything is written. This surfaces parse errors, suspicious overwrites and mapping issues early, on the full dataset, instead of halfway through a write. Apply copies the input gpkg into `output/` (as `<name>_updated_<timestamp>.gpkg`) and executes exactly the reviewed plan on that copy inside a single transaction (rollback on any error), finishing by bumping `last_change` in `gpkg_contents`. The input gpkg is never modified, so it doubles as the backup.

**Per-column kill switches.** `APPLY_COLUMNS` at the top of the script enables/disables each mapping individually, so uncertain mappings (e.g. `opp_pan`, `le_vorm`) can be vetoed without code changes. `jr_vrea` ships disabled.

**Empty never beats filled.** In this monitor, `0` is the generic "not filled in" value for nearly every column. The script therefore refuses to overwrite a filled gpkg value with NULL; such cases become warnings instead of mutations. Columns where clearing _is_ desired can be whitelisted in `ALLOW_CLEAR`. Two deliberate exceptions to the 0-handling:

- `Lokaal eigendom = 0` **is** meaningful (no local ownership) and maps to `Nee`;
- literal `'0'` strings found _inside gpkg text fields_ (a data quirk, e.g. in `zv_subs_id`) are treated as empty on the gpkg side, so real values can replace them without noise.

**Stopped projects are deleted.** A row with fase `Project gestopt` that still exists in the gpkg generates a `DELETE` mutation — removing every row that shares its project id (relevant for multi-row wind projects) — and is reported as `STOPPED_DELETED`; stopped projects absent from the gpkg are `STOPPED_SKIPPED`. This is governed by the `DELETE_STOPPED` switch at the top of the script; set it to `False` to revert to report-only (`STOPPED_STILL_IN_GPKG`, no deletion).

**New projects are inserted, with placeholder geometry.** Monitor rows without a gpkg counterpart (and not stopped) are inserted with all mapped attributes; wind rows get `wl_tur_id = ES05_{nr}_totaal` following the existing convention. Geometry: a 100 × 100 m square centred on the monitor's X/Y coordinate (RD New / EPSG:28992), written as a hand-assembled GeoPackage binary blob (GP header incl. envelope + WKB MultiPolygon) — enough to make the feature visible in QGIS for proper digitising. Rows without plausible RD coordinates get NULL geometry and a per-project note in the coverage report.

**Plain-sqlite inserts vs. GDAL's rtree triggers.** GeoPackages written by GDAL carry spatial-index triggers that call SQL functions like `ST_IsEmpty` and `ST_MinX` — functions that only exist when GDAL registers them at runtime. Under plain `sqlite3`, any INSERT/UPDATE/DELETE would fail at statement-prepare time. The script solves this by registering minimal Python implementations of those functions (via `Connection.create_function`) that read the envelope straight from the geometry blob header. Verified after apply: the rtree entry count matches the non-NULL geometry count, and GDAL's `gpkg_ogr_contents` feature-count triggers fire correctly.

**Coverage as a first-class output.** Every project ends up in exactly one bucket — matched, `NEW_IN_EXCEL`, `ONLY_IN_GPKG`, `STOPPED_SKIPPED`, or `STOPPED_DELETED` (`STOPPED_STILL_IN_GPKG` when `DELETE_STOPPED` is off) — so gaps between the two sources are always explicit rather than silently absorbed.

## Known data quirks

Documented here because they will bite again:

- The Excel header of the gross-production column is `"Productie "` — with a trailing space. The script normalises headers on load. Do not map this column to `*_pr`: gross production excludes the realisatiegraad; `Prognose_GWh` is the correct source for `*_pr`.
- Onderstation spelling drift: monitor "Zwolle Wetering**s**kade" vs gpkg "Zwolle Weteringkade" (auto-corrected). Note that NP RES now asks for voedingsgebied-IDs from the Netbeheer Nederland CSV rather than free-text names — both sources still use names, so that migration is a separate future task.

## Limitations / manual follow-up

- Geometry is never _updated_ for existing projects — the monitor only has a point coordinate, the gpkg needs polygons. Placeholder squares and NULL-geometry inserts must be digitised in QGIS.
- The `le_vorm` code-4 mapping remains a judgement call, flagged in warnings.
- **`opp_pan` — resolved 2026-09-24: do not sync.** The two sources measure different things. The monitor's `Oppervlak zonnepark` (which the script copies, `update_res_gpkg.py:358`) is the **area of the whole park**; NP RES asks for the **area of the PV panels**, which is what the aanlevering gpkg holds — derived in an earlier round from the panel count, hence its long decimals (`12.3456790123457` vs the monitor's `6.7`). The gpkg values are leading. Accept that newly inserted projects have a gap here until the panel area is worked out by hand.
- Stopped-project deletions are executed on apply (see the `DELETE_STOPPED` switch). Obsolete gpkg-only rows (`ONLY_IN_GPKG`) are still only reported, never deleted.

### Open digitising queue — carry forward to the next update

All rows now have a geometry (2026-10-19), but these seven carry a **placeholder square** rather than a real footprint — legible on a map, not an actual delineation. Replace with real footprints once locations are known. Convention: wind placeholder = 25 × 25 m square (individual turbines not yet known), solar placeholder = 100 × 100 m square.

| Layer | Project | Naam | Gemeente | Placeholder centre (RD) | Source of location |
| ----- | ------- | ---- | -------- | ----------------------- | ------------------ |
| `npres_wol_v4` | `ES05_94`  | Broeklanden | Hardenberg | 239591, 508168 | coordinates supplied by user |
| `npres_wol_v4` | `ES05_119` | Kloosterhaar | Hardenberg | 242984, 502996 | village centroid (PDOK) — rough, not a real turbine site |
| `npres_wol_v4` | `ES05_123` | Windpark Groen | Steenwijkerland | 198395, 529995 | earlier placeholder (was 10 × 10 m) |
| `npres_zov_v4` | `ES05_120` | Zonneveld vd Valk | Zwolle | 207331, 505283 | Van der Valk Hotel, Nieuwleusenerdijk 1 (~3 ha park next to the hotel) |
| `npres_zov_v4` | `ES05_121` | Zonneveld A28 (OER project) | Zwolle | 208276, 505397 | Hessenpoort centroid — rough; actual candidate sites are a strip by Hessenpoort and the A28 Ommen-exit slope (~1.86 ha total) |
| `npres_zov_v4` | `ES05_122` | Eeser Campus | Steenwijkerland | 205210, 535253 | monitor X/Y (address Eeser Campus 30 is ~820 m away, at 205843, 534724) |
| `npres_zov_v4` | `ES05_37`  | Zonnepark en dak Vollenhove waterschap | Steenwijkerland | 192949, 520583 | monitor X/Y, ~70 m from the address Noordwal 7, 8325PP Vollenhove (WDOD). Added 2026-10-20: the old geometry was a 2 × 2 m marker at 208686, 488648 near Raalterweg, Wijhe (wrong municipality, 35.6 km off), present since `projecten_geopackage_2025_q4` |

To find placeholders at any time — NULL geometry, or a bounding box of exactly 25 m (wind) / 100 m (solar) wide in the rtree:

```sql
SELECT t.wl_prj_id, t.wl_naam_p FROM npres_wol_v4 t LEFT JOIN rtree_npres_wol_v4_geom r ON r.id = t.fid
 WHERE t.geom IS NULL OR ROUND(r.maxx - r.minx, 1) = 25.0 AND ROUND(r.maxy - r.miny, 1) = 25.0
UNION ALL
SELECT t.zv_prj_id, t.zv_naam FROM npres_zov_v4 t LEFT JOIN rtree_npres_zov_v4_geom r ON r.id = t.fid
 WHERE t.geom IS NULL OR ROUND(r.maxx - r.minx, 1) = 100.0 AND ROUND(r.maxy - r.miny, 1) = 100.0;
```

Note for the wind layer: a digitised project should end up with **one row per turbine**, and `wl_verm` / `wl_pr` are then **per turbine**, not park totals (rows are summed downstream). ES05_94 currently holds 12 MW on a single row for `wl_aant_t` = 2 — splitting it into two turbines means 6 MW each.

## `zv_trede` suggestions — FOR MANUAL REVIEW (not applied)

Suggested `zv_trede` values for Zon-op-Veld projects that currently have no value, derived by overlaying `npres_zov_v4` with the Tredenkaart (`NPRES Dataportaal/Tredenkaart_ES05/Tredenkaart_ES05.gpkg`). Method: take each project's guaranteed-inside center point (`ST_PointOnSurface`) and test which trede-layer polygon contains it, checking `trede4 → trede3 → trede2`, first match wins. `trede1_gebouwen` is not tested (per-building layer, irrelevant for field solar). Both layers are EPSG:28992; no reprojection.

**⚠️ These are suggestions only — nothing was written to any GeoPackage. Review before use.** When the same method is validated against the 44 projects that *already* have a `zv_trede`, it only reproduces the existing value in **25 of 44 cases (57%)**, disagreeing in both directions. So the automated overlay is not reliable enough to trust blindly; treat the values below as a starting point, especially the ambiguous rows.

| zv_prj_id | Project | Gemeente | Suggested | Trede layers the center falls in |
| --------- | ------- | -------- | --------- | -------------------------------- |
| ES05_76   | Onder de Gelder            | GM1773 | Trede 4 | 4 (only)      |
| ES05_88   | Blokzijl "De Hoop"         | GM1708 | Trede 4 | 4 (only)      |
| ES05_104  | Energielandgoed Raalte     | GM0177 | Trede 4 | 4 (only)      |
| ES05_109  | Zonnepark Marshoek         | GM0148 | Trede 4 | 4 (only)      |
| ES05_112  | Zonnepark Holterweg        | GM0150 | Trede 4 | 4 (only)      |
| ES05_113  | Zonnepark Bussinksweg      | GM0150 | Trede 4 | 4 (only)      |
| ES05_116  | Zandwinplas Hasselt        | GM1896 | Trede 3 | 3 (only)      |
| ES05_122  | Eeser Campus               | GM1708 | Trede 2 | 2 (only)      |
| ES05_120  | Zonneveld vd Valk          | GM0193 | Trede 4 ⚠️ | **2, 3 and 4** — ambiguous, rule picked 4 |
| ES05_121  | Zonneveld A28 (OER project)| GM0193 | Trede 4 ⚠️ | **2, 3 and 4** — ambiguous, rule picked 4 |

The first 8 projects fall inside exactly one trede layer, so those suggestions are unambiguous (subject to the overall 57% caveat). **ES05_120** and **ES05_121** fall inside all three of trede2/3/4 simultaneously — the 4→3→2 order picked `Trede 4`, but `Trede 2` or `Trede 3` are equally defensible; decide these by hand.
