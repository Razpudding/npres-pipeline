# RES West-Overijssel — Monitor → GeoPackage sync

Tooling to update the NP RES _projecten-geopackage_ (layers `npres_zov_v4` — Zon op Veld, and `npres_wol_v4` — Wind op Land) from the regional monitor Excel of RES West-Overijssel.

The RES region tracks solar-field and onshore-wind projects in an Excel monitor with its own column names, value conventions and phase model. NP RES requires the same data as a GeoPackage that follows the national _invulformat_ data model (fixed column names, field lengths and choice lists). This repo bridges the two: it documents the column mapping between both models and provides a script that applies monitor updates to the GeoPackage in a controlled, reviewable way.

## Contents

| File                                    | Purpose                                                                                                                                                                                                                                                                                                               |
| --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `update_res_gpkg.py`                    | Two-phase update script (plan → apply)                                                                                                                                                                                                                                                                                |
| `kolom_mapping_res_monitor.xlsx`        | Column mapping workbook: per gpkg column the matched Excel column, match confidence, required transformation, and comments. Also contains ID-coverage analysis and a value/format differences report between the two sources. Editable — this is the source of truth for the mapping decisions encoded in the script. |
| `NP_RES_data_portaal_instructies_*.pdf` | NP RES documentation of the GeoPackage data model (attribute descriptions per layer)                                                                                                                                                                                                                                  |

## Usage

```bash
# dry run: compute and report all mutations, change nothing
python update_res_gpkg.py --gpkg projecten.gpkg --excel monitor.xlsx

# execute (after reviewing the plan); asks for confirmation, makes a backup first
python update_res_gpkg.py --gpkg projecten.gpkg --excel monitor.xlsx --apply

# the monitor sheet is auto-detected (first sheet starting with "Monitor");
# override with --sheet "Monitor Q2 2026"
```

Requirements: Python 3.9+, `pandas`, `openpyxl`. **No GDAL/GeoPandas needed** — a GeoPackage is a SQLite database, and the script talks to it directly via the stdlib `sqlite3` module.

Every plan run writes three CSVs next to the gpkg for review:

- `plan_mutations.csv` — every intended change: layer, fid, project id, column, old value, new value
- `plan_warnings.csv` — everything that needs a human decision (unparseable values, refused overwrites, uncertain mappings)
- `plan_coverage.csv` — per project: matched / new in Excel / only in gpkg / stopped

## How the two datasets relate

- **Join key:** Excel column `Nr.` (integer) ↔ gpkg `zv_prj_id` / `wl_prj_id` as `"ES05_" + Nr`. Leading zeros in Excel are display formatting only, so `Nr` is cast to `int` before concatenating. The join is per layer: Excel `Type` routes rows to `npres_zov_v4` (`Zon-op-veld`) or `npres_wol_v4` (`Wind`); `Zon-op-dak` rows are out of scope. `Nr` values may repeat _across_ types (e.g. Nr 116 exists as both a solar and a wind project), which is safe because the effective key is `Nr + Type`.
- **Snapshot difference:** the gpkg represents an earlier reporting round (Q4 2025); the monitor is current (Q2 2026). Most value differences are therefore legitimate progress: phases advanced, Slaagkans changed, production figures updated.
- **The Excel is leading** for all mapped attribute columns. The gpkg remains leading for everything the monitor doesn't track (geometry, trede, opslag, transportcapaciteit, …) and for a small set of "special case" rows that intentionally exist only in the gpkg — those are reported but never touched.

## Column mapping (summary)

The full mapping with per-column rationale lives in `kolom_mapping_res_monitor.xlsx`. The short version:

**Direct copies:** `Projectnaam→*_naam`, `Vermogen→*_verm`, `Slaagkans→*_reagr_r`, `Prognose_GWh→*_pr` (its Excel formula, Slaagkans × Productie, is exactly the NP RES definition of `pr`), `SDE→*_subs_id`, `Onderstation→*_ondsta(t)`, `Aantal turbines→wl_aant_t`, `Ashoogte→wl_masth`, `Tiphoogte→wl_tiph`, `Oppervlak zonnepark→zv_opp_pan`.

**Value maps / transformations:**

- `Fase`: the monitor's labels are mapped onto the NP RES choice list, absorbing the monitor's spelling quirks (`Productie→huidig`, `Vergunningsaanvraag→vergunningaanvraag`, `Subsidiebeschiking en bouw` [sic, one *k*] `→subsidiebeschikking en bouw`).
- `Gemeente` (name) → CBS `GM####` code via a hardcoded lookup of the eleven West-Overijssel municipalities, verified empirically against the gpkg.
- `F3/F4/F6` → `*_dt_va` / `*_dt_vv` / `*_dt_rea`. The F-columns are "date status _n_ started" per the monitor's own column spec; F3/F4/F6 correspond to the vergunningaanvraag/vergunningverlening/productie phases. Month-year precision, stored as ISO dates.
- `Lokaal eigendom`: a fraction in Excel (with pollution like `"100%"`, `"25%"`, `" -"`), a 0–100 percentage in the gpkg. Parsed defensively, ×100, and additionally derived into the `*_le` Ja/Nee flag.
- `Vorm lokaal eigendom`: monitor digits → NP RES letter codes. **The mapping is not positional**: monitor 1 = financiële participatie zonder eigendom → `A (ze)`, 2 = omgevingsfonds → `C (fon)`, 3 = omwonendenregeling → `B (omw)`, and 4 = procesparticipatie has no clean NP RES equivalent — it is mapped to `D (and)` and every occurrence emits a _verify_ warning. Combination strings (`"2,3,4"`, `"1 en 4"`) are parsed and rendered as the corresponding NP RES combination label (`BCD (omw+fon+and)`).
- `rescode` is the constant `ES05`.

**Derived (no Excel source column):** `*_vol_r = round(Productie / Vermogen × 1000)` and `*_pr_plam = Productie × (1 − Slaagkans)`, both validated against the existing gpkg values before adopting the formulas.

**Deliberately unmapped:** `zv_trede` (the monitor's "Type zonnepark" values were tested against it — no correlation; they encode the provincial zonneladder, not the national voorkeursvolgorde), `*_bev_gez` (the monitor's "Omgevingsvergunning" is a Ja/Nee _granted_ flag, not the bestuurslaag), `zv_jr_vrea` (source column is nearly empty and contains junk like Excel-serial dates and `"22/23"` — disabled by default), and everything the monitor simply doesn't track (`tra_cap`, `opslag`, `rol`, `bgt_id`, `brt_id`, `verm_om`, `le_part`).

## Design decisions in the script

**Two-phase execution.** The plan phase computes every mutation against in-memory copies of both sources and reports them (counts per column, warnings, coverage) before anything is written. This surfaces parse errors, suspicious overwrites and mapping issues early, on the full dataset, instead of halfway through a write. Apply executes exactly the reviewed plan inside a single transaction (rollback on any error) after taking a timestamped backup of the gpkg, and finishes by bumping `last_change` in `gpkg_contents`.

**Per-column kill switches.** `APPLY_COLUMNS` at the top of the script enables/disables each mapping individually, so uncertain mappings (e.g. `opp_pan`, `le_vorm`) can be vetoed without code changes. `jr_vrea` ships disabled.

**Empty never beats filled.** In this monitor, `0` is the generic "not filled in" value for nearly every column. The script therefore refuses to overwrite a filled gpkg value with NULL; such cases become warnings instead of mutations. Columns where clearing _is_ desired can be whitelisted in `ALLOW_CLEAR`. Two deliberate exceptions to the 0-handling:

- `Lokaal eigendom = 0` **is** meaningful (no local ownership) and maps to `Nee`;
- literal `'0'` strings found _inside gpkg text fields_ (a data quirk, e.g. in `zv_subs_id`) are treated as empty on the gpkg side, so real values can replace them without noise.

**Stopped projects are skipped, not deleted.** Rows with fase `Project gestopt` generate no mutations. If a stopped project still exists in the gpkg it is flagged (`STOPPED_STILL_IN_GPKG`) for a manual removal decision — the script never deletes features.

**New projects are inserted, with placeholder geometry.** Monitor rows without a gpkg counterpart (and not stopped) are inserted with all mapped attributes; wind rows get `wl_tur_id = ES05_{nr}_totaal` following the existing convention. Geometry: a 100 × 100 m square centred on the monitor's X/Y coordinate (RD New / EPSG:28992), written as a hand-assembled GeoPackage binary blob (GP header incl. envelope + WKB MultiPolygon) — enough to make the feature visible in QGIS for proper digitising. Rows without plausible RD coordinates get NULL geometry and a per-project note in the coverage report.

**Plain-sqlite inserts vs. GDAL's rtree triggers.** GeoPackages written by GDAL carry spatial-index triggers that call SQL functions like `ST_IsEmpty` and `ST_MinX` — functions that only exist when GDAL registers them at runtime. Under plain `sqlite3`, any INSERT would fail at statement-prepare time. The script solves this by registering minimal Python implementations of those functions (via `Connection.create_function`) that read the envelope straight from the geometry blob header. Verified after apply: the rtree entry count matches the non-NULL geometry count, and GDAL's `gpkg_ogr_contents` feature-count triggers fire correctly.

**Coverage as a first-class output.** Every project ends up in exactly one bucket — matched, `NEW_IN_EXCEL`, `ONLY_IN_GPKG`, `STOPPED_SKIPPED`, or `STOPPED_STILL_IN_GPKG` — so gaps between the two sources are always explicit rather than silently absorbed.

## Known data quirks

Documented here because they will bite again:

- The Excel header of the gross-production column is `"Productie "` — with a trailing space. The script normalises headers on load. Do not map this column to `*_pr`: gross production excludes the realisatiegraad; `Prognose_GWh` is the correct source for `*_pr`.
- Onderstation spelling drift: monitor "Zwolle Wetering**s**kade" vs gpkg "Zwolle Weteringkade" (auto-corrected). Note that NP RES now asks for voedingsgebied-IDs from the Netbeheer Nederland CSV rather than free-text names — both sources still use names, so that migration is a separate future task.

## Limitations / manual follow-up

- Geometry is never _updated_ for existing projects — the monitor only has a point coordinate, the gpkg needs polygons. Placeholder squares and NULL-geometry inserts must be digitised in QGIS.
- The `le_vorm` code-4 mapping and the `opp_pan` semantics (panel area vs. park area) remain judgement calls; both are flagged in warnings / the mapping workbook.
- Deletions (stopped projects, obsolete gpkg-only rows) are reported, never executed.
