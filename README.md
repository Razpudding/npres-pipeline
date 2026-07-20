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

**Deliberately unmapped:** `zv_trede` (the monitor's "Type zonnepark" values were tested against it — no correlation; they encode the provincial zonneladder, not the national voorkeursvolgorde), `*_bev_gez` (the monitor's "Omgevingsvergunning" is a Ja/Nee _granted_ flag, not the bestuurslaag), `zv_jr_vrea` (source column is nearly empty and contains junk like Excel-serial dates and `"22/23"` — disabled by default), and everything the monitor simply doesn't track (`tra_cap`, `opslag`, `rol`, `bgt_id`, `brt_id`, `verm_om`, `le_part`).

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
- The `le_vorm` code-4 mapping and the `opp_pan` semantics (panel area vs. park area) remain judgement calls; both are flagged in warnings.
- Stopped-project deletions are executed on apply (see the `DELETE_STOPPED` switch). Obsolete gpkg-only rows (`ONLY_IN_GPKG`) are still only reported, never deleted.

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
