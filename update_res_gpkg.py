#!/usr/bin/env python3
"""
Update npres_zov_v4 / npres_wol_v4 layers in the RES projecten-geopackage
from the RES West monitor Excel.

Two-phase:
  1. PLAN  (default)  - reads both files, computes every mutation, prints a
     summary (mutations per column, inserts, warnings, coverage) and writes
     plan_mutations.csv / plan_warnings.csv / plan_coverage.csv next to the gpkg.
  2. APPLY (--apply)  - makes a timestamped backup of the gpkg, then executes
     the exact mutations from the plan (attribute UPDATEs + INSERTs for new
     projects with a placeholder square geometry at the monitor X/Y).

Files are read from the input/ folder and results (plan CSVs + updated gpkg)
are written to the output/ folder. The input files are never modified.

Usage:
    # put projecten.gpkg and monitor.xlsx in input/, then:
    python update_res_gpkg.py --gpkg projecten.gpkg --excel monitor.xlsx [--sheet "Monitor Q2 2026"] [--apply]

Requires: pandas, openpyxl. Pure sqlite3 for the gpkg (no GDAL needed).

Transformation logic follows kolom_mapping_res_monitor.xlsx. Per-column
switches live in APPLY_COLUMNS below, so you can veto individual mappings
(e.g. the uncertain ones) without touching the code.
"""

import argparse
import csv
import math
import re
import shutil
import sqlite3
import struct
import sys
from datetime import datetime, date, timezone
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# Default folders (relative to the current working directory, i.e. run from the
# project root). Input files are read from here and never modified; the updated
# gpkg and all plan CSVs are written to the output dir.
INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")

RESCODE = "ES05"

# Per-column switch: set to False to exclude a mapping from the update.
# Keys are gpkg column names (both layers share the switch via suffix).
APPLY_COLUMNS = {
    "naam": True,          # zv_naam / wl_naam_p <- Projectnaam
    "fase": True,          # <- Fase (value map)
    "gemcode": True,       # <- Gemeente (name -> GM-code)
    "dt_va": True,         # <- F3 (start fase vergunningsaanvraag)
    "dt_vv": True,         # <- F4 (start fase vergunningverlening)
    "dt_rea": True,        # <- F6 (start fase productie)
    "jr_vrea": False,      # <- 'Verwachte realisatie': source column is unreliable, off by default
    "subs_id": True,       # <- SDE
    "opp_pan": True,       # <- Oppervlak zonnepark (ZoV only, uncertain semantics)
    "verm": True,          # <- Vermogen
    "vol_r": True,         # derived: Productie / Vermogen * 1000
    "reagr_r": True,       # <- Slaagkans
    "pr": True,            # <- Prognose_GWh
    "pr_plam": True,       # derived: Productie * (1 - Slaagkans)
    "ondsta": True,        # zv_ondsta / wl_ondstat <- Onderstation
    "rescode": True,       # constant ES05
    "le": True,            # derived from Lokaal eigendom > 0
    "le_p": True,          # <- Lokaal eigendom (fraction -> 0-100)
    "le_vorm": True,       # <- Vorm lokaal eigendom (digits -> NP RES letter codes, uncertain)
    "aant_t": True,        # WoL <- Aantal turbines
    "masth": True,         # WoL <- Ashoogte
    "tiph": True,          # WoL <- Tiphoogte
}

# Never overwrite a filled gpkg value with NULL (Excel 0/blank = "unknown").
# Columns listed here MAY be cleared by an empty Excel value:
ALLOW_CLEAR = set()  # e.g. {"zv_ondsta"} if you want empties to propagate

FASE_MAP = {
    "productie": "huidig",
    "vergunningsaanvraag": "vergunningaanvraag",
    "vergunningaanvraag": "vergunningaanvraag",
    "vergunningverlening": "vergunningverlening",
    "subsidiebeschiking en bouw": "subsidiebeschikking en bouw",
    "subsidiebeschikking en bouw": "subsidiebeschikking en bouw",
    "voortraject": "voortraject",
}
FASE_STOPPED = "project gestopt"

GEMCODE = {
    "dalfsen": "GM0148", "deventer": "GM0150", "hardenberg": "GM0160",
    "kampen": "GM0166", "olst-wijhe": "GM1773", "ommen": "GM0175",
    "raalte": "GM0177", "staphorst": "GM0180", "steenwijkerland": "GM1708",
    "zwartewaterland": "GM1896", "zwolle": "GM0193",
}

ONDSTA_FIX = {"zwolle weteringskade": "Zwolle Weteringkade"}

# Excel 'Vorm lokaal eigendom' legend -> NP RES letter codes:
#   1 = financiele participatie zonder eigendom -> A (ze)
#   2 = omgevingsfonds                          -> C (fon)
#   3 = omwonendenregeling                      -> B (omw)
#   4 = procesparticipatie -> closest is D (and) "andere lokale baten"  [UNCERTAIN]
LE_VORM_DIGIT = {"1": "A", "2": "C", "3": "B", "4": "D"}
LE_VORM_LABEL = {
    "A": "A (ze)", "B": "B (omw)", "C": "C (fon)", "D": "D (and)",
    "AB": "AB (ze+omw)", "AC": "AC (ze+fon)", "AD": "AD (ze+and)",
    "BC": "BC (omw+fon)", "BD": "BD (omw+and)", "CD": "CD (fon+and)",
    "ABC": "ABC (ze+omw+fon)", "ACD": "ACD (ze+fon+and)",
    "BCD": "BCD (omw+fon+and)", "ABCD": "ABCD (ze+omw+fon+and)",
}

PLACEHOLDER_HALF_SIZE = 50.0  # meters; placeholder square = 100 x 100 m around X/Y


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

def is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    s = str(v).strip()
    return s in ("", "0", "0.0", "nan", "None", "-")


def to_float(v):
    if is_blank(v):
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def clean_text(v):
    if is_blank(v):
        return None
    return str(v).strip()


def to_iso_date(v):
    if is_blank(v):
        return None
    if isinstance(v, (datetime, date)):
        if getattr(v, "year", 0) < 1990:  # Excel artefacts like 1905-07-14
            return None
        return v.strftime("%Y-%m-%d")
    return None  # anything non-datetime ('22/23' etc.) is rejected


def parse_le_fraction(v):
    """'Lokaal eigendom': fraction 0-1, or '100%' / '25%' strings. Returns 0-1 or None.
    Note: 0 is meaningful here (= no local ownership), unlike other monitor columns."""
    if v is not None and str(v).strip() in ("0", "0.0"):
        return 0.0
    if is_blank(v):
        return None
    s = str(v).strip()
    m = re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*%", s)
    if m:
        return float(m.group(1).replace(",", ".")) / 100.0
    f = to_float(s)
    if f is None:
        return None
    if f > 1.0:  # already a percentage typed as number
        return f / 100.0
    return f


def parse_le_vorm(v, warn):
    if is_blank(v):
        return None
    s = str(v).strip().lower()
    digits = sorted({LE_VORM_DIGIT[d] for d in re.findall(r"[1-4]", s)})
    if not digits:
        warn(f"le_vorm value '{v}' contains no codes 1-4 -> skipped")
        return None
    if "D" in digits:
        warn(f"le_vorm '{v}': Excel code 4 (procesparticipatie) mapped to D (and) - verify")
    return LE_VORM_LABEL["".join(digits)]


# --------------------------------------------------------------------------
# GeoPackage geometry blob (placeholder square, EPSG:28992)
# --------------------------------------------------------------------------

def square_multipolygon_blob(x, y, half=PLACEHOLDER_HALF_SIZE, srs=28992):
    pts = [(x - half, y - half), (x + half, y - half), (x + half, y + half),
           (x - half, y + half), (x - half, y - half)]
    wkb = struct.pack("<BI", 1, 6)              # little endian, MultiPolygon
    wkb += struct.pack("<I", 1)                 # 1 polygon
    wkb += struct.pack("<BI", 1, 3)             # Polygon
    wkb += struct.pack("<I", 1)                 # 1 ring
    wkb += struct.pack("<I", len(pts))
    for px, py in pts:
        wkb += struct.pack("<dd", px, py)
    header = b"GP" + struct.pack("<BB", 0, 0x03)  # version 0, LE + XY envelope
    header += struct.pack("<i", srs)
    header += struct.pack("<dddd", x - half, x + half, y - half, y + half)
    return header + wkb


def valid_rd(x, y):
    return x is not None and y is not None and 0 < x < 300000 and 289000 < y < 630000


def register_st_functions(con):
    """Minimal ST_* implementations so GDAL's rtree triggers work under plain sqlite3."""
    def envelope(blob):
        if blob is None:
            return None
        flags = blob[3]
        env_ind = (flags >> 1) & 0x07
        if env_ind == 0:
            return None
        return struct.unpack_from("<dddd", blob, 8)  # minx, maxx, miny, maxy

    con.create_function("ST_IsEmpty", 1, lambda b: 1 if b is None else 0)
    con.create_function("ST_MinX", 1, lambda b: envelope(b)[0] if envelope(b) else None)
    con.create_function("ST_MaxX", 1, lambda b: envelope(b)[1] if envelope(b) else None)
    con.create_function("ST_MinY", 1, lambda b: envelope(b)[2] if envelope(b) else None)
    con.create_function("ST_MaxY", 1, lambda b: envelope(b)[3] if envelope(b) else None)


# --------------------------------------------------------------------------
# Excel parsing
# --------------------------------------------------------------------------

def load_monitor(path, sheet):
    xls = pd.ExcelFile(path)
    if sheet is None:
        candidates = [s for s in xls.sheet_names if s.lower().startswith("monitor")]
        if not candidates:
            sys.exit(f"No sheet starting with 'Monitor' found. Sheets: {xls.sheet_names}")
        sheet = candidates[0]
    raw = pd.read_excel(path, sheet_name=sheet, header=None)

    header_row = None
    for i in range(min(10, len(raw))):
        vals = raw.iloc[i].astype(str).str.strip().tolist()
        if "Nr." in vals and "Projectnaam" in vals:
            header_row = i
            break
    if header_row is None:
        sys.exit("Could not locate the header row (expected 'Nr.' and 'Projectnaam').")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row]]
    df = df[df["Type"].isin(["Zon-op-veld", "Wind"])]
    df = df[df["Nr."].notna()]
    df = df[df["Nr."].map(lambda v: isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)))]
    df["gid"] = RESCODE + "_" + df["Nr."].astype(int).astype(str)
    print(f"Excel sheet '{sheet}': {len(df)} veld/wind rows "
          f"({(df['Type'] == 'Zon-op-veld').sum()} zon-op-veld, {(df['Type'] == 'Wind').sum()} wind)")
    dup = df[df.duplicated(subset=["gid", "Type"], keep=False)]
    if len(dup):
        print("WARNING: duplicate Nr within same Type:\n", dup[["gid", "Type", "Projectnaam"]].to_string())
    return df, sheet


# --------------------------------------------------------------------------
# Row -> target values (returns dict of gpkg_column -> value, plus warnings)
# --------------------------------------------------------------------------

def transform_row(r, layer, warn):
    """layer: 'zov' or 'wol'. Returns dict gpkg_col -> new value (only enabled columns)."""
    p = "zv_" if layer == "zov" else "wl_"
    out = {}

    def put(short, col, val):
        if APPLY_COLUMNS.get(short, False):
            out[col] = val

    naam = clean_text(r.get("Projectnaam"))
    put("naam", p + ("naam" if layer == "zov" else "naam_p"), naam)

    fase_raw = str(r.get("Fase", "")).strip().lower()
    fase = FASE_MAP.get(fase_raw)
    if fase is None and fase_raw != FASE_STOPPED:
        warn(f"unknown fase '{r.get('Fase')}' -> fase not updated")
    put("fase", p + "fase", fase)

    gem = clean_text(r.get("Gemeente"))
    if gem:
        code = GEMCODE.get(gem.lower())
        if code is None:
            warn(f"unknown gemeente '{gem}' -> gemcode not updated")
        put("gemcode", p + "gemcode", code)

    put("dt_va", p + "dt_va", to_iso_date(r.get("F3")))
    put("dt_vv", p + "dt_vv", to_iso_date(r.get("F4")))
    put("dt_rea", p + "dt_rea", to_iso_date(r.get("F6")))

    if APPLY_COLUMNS.get("jr_vrea"):
        vr = r.get("Verwachte realisatie")
        yr = None
        if isinstance(vr, (datetime, date)) and vr.year > 1990:
            yr = vr.year
        elif to_float(vr) and 1990 < to_float(vr) < 2100:
            yr = int(to_float(vr))
        elif not is_blank(vr):
            warn(f"unparseable 'Verwachte realisatie' '{vr}'")
        out[p + "jr_vrea"] = yr

    sde = clean_text(r.get("SDE"))
    if sde is not None:
        low = sde.lower()
        sde = "Ja" if low == "ja" else "Nee" if low == "nee" else sde
    put("subs_id", p + "subs_id", sde)

    verm = to_float(r.get("Vermogen"))
    prod = to_float(r.get("Productie"))  # header has trailing space; normalized on load
    slaag = to_float(r.get("Slaagkans"))
    if slaag is None and not is_blank(r.get("Slaagkans")):
        warn(f"unparseable Slaagkans '{r.get('Slaagkans')}'")
    prog = to_float(r.get("Prognose_GWh"))

    put("verm", p + "verm", verm)
    # Slaagkans 0 only occurs for stopped projects, which are skipped upstream.
    put("reagr_r", p + "reagr_r", slaag)
    put("pr", p + "pr", prog)
    if APPLY_COLUMNS.get("vol_r"):
        out[p + "vol_r"] = round(prod / verm * 1000) if (prod and verm) else None
    if APPLY_COLUMNS.get("pr_plam"):
        out[p + "pr_plam"] = round(prod * (1 - slaag), 4) if (prod is not None and slaag is not None) else None

    ond = clean_text(r.get("Onderstation"))
    if ond:
        ond = ONDSTA_FIX.get(ond.lower(), ond)
    put("ondsta", p + ("ondsta" if layer == "zov" else "ondstat"), ond)

    put("rescode", "rescode", RESCODE)

    le_frac = parse_le_fraction(r.get("Lokaal eigendom"))
    if le_frac is None and not is_blank(r.get("Lokaal eigendom")):
        warn(f"unparseable 'Lokaal eigendom' '{r.get('Lokaal eigendom')}'")
    put("le", p + "le", None if le_frac is None else ("Ja" if le_frac > 0 else "Nee"))
    put("le_p", p + "le_p", round(le_frac * 100, 2) if le_frac else None)
    put("le_vorm", p + "le_vorm", parse_le_vorm(r.get("Vorm lokaal eigendom"), warn))

    if layer == "zov":
        put("opp_pan", "zv_opp_pan", to_float(r.get("Oppervlak zonnepark")))
    else:
        aant = to_float(r.get("Aantal turbines"))
        put("aant_t", "wl_aant_t", int(aant) if aant else None)
        put("masth", "wl_masth", to_float(r.get("Ashoogte")))
        put("tiph", "wl_tiph", to_float(r.get("Tiphoogte")))

    return out


# --------------------------------------------------------------------------
# Plan
# --------------------------------------------------------------------------

LAYERS = {
    "zov": {"table": "npres_zov_v4", "id_col": "zv_prj_id", "excel_type": "Zon-op-veld"},
    "wol": {"table": "npres_wol_v4", "id_col": "wl_prj_id", "excel_type": "Wind"},
}


def values_equal(old, new):
    if old is None and new is None:
        return True
    if isinstance(old, (int, float)) and isinstance(new, (int, float)):
        return math.isclose(float(old), float(new), rel_tol=1e-6, abs_tol=1e-6)
    return str(old) == str(new)


def build_plan(con, excel_df):
    mutations, warnings, coverage = [], [], []

    for layer, cfg in LAYERS.items():
        table, id_col = cfg["table"], cfg["id_col"]
        gpkg = pd.read_sql(f'SELECT * FROM "{table}"', con)
        gpkg_ids = dict(zip(gpkg[id_col], gpkg["fid"]))
        gpkg_rows = gpkg.set_index("fid").to_dict("index")

        exl = excel_df[excel_df["Type"] == cfg["excel_type"]]
        excel_ids = set(exl["gid"])

        for _, r in exl.iterrows():
            gid = r["gid"]
            stopped = str(r.get("Fase", "")).strip().lower() == FASE_STOPPED
            row_warn = lambda msg, g=gid: warnings.append([layer, g, r.get("Projectnaam"), msg])

            if stopped:
                if gid in gpkg_ids:
                    coverage.append([layer, gid, r.get("Projectnaam"),
                                     "STOPPED_STILL_IN_GPKG",
                                     "Excel says 'Project gestopt' but project exists in gpkg - remove manually if desired"])
                else:
                    coverage.append([layer, gid, r.get("Projectnaam"), "STOPPED_SKIPPED", ""])
                continue

            new_vals = transform_row(r, layer, row_warn)

            if gid in gpkg_ids:
                fid = gpkg_ids[gid]
                old = gpkg_rows[fid]
                for col, new in new_vals.items():
                    oldv = old.get(col)
                    if isinstance(oldv, float) and math.isnan(oldv):
                        oldv = None
                    if isinstance(oldv, str) and oldv.strip() == "0":
                        oldv = None  # gpkg contains literal '0' placeholders in some text fields
                    if values_equal(oldv, new):
                        continue
                    if new is None and oldv is not None and col not in ALLOW_CLEAR:
                        warnings.append([layer, gid, r.get("Projectnaam"),
                                         f"{col}: Excel empty/0 but gpkg has '{oldv}' -> kept gpkg value (add to ALLOW_CLEAR to clear)"])
                        continue
                    mutations.append([layer, "UPDATE", gid, fid, col, oldv, new])
            else:
                x, y = to_float(r.get("X_coordinaat")), to_float(r.get("Y_coordinaat"))
                geom_note = "placeholder 100m square at X/Y" if valid_rd(x, y) else "NO valid X/Y -> NULL geometry, digitize manually"
                ins = dict(new_vals)
                ins[id_col] = gid
                if layer == "wol":
                    ins["wl_tur_id"] = f"{gid}_totaal"
                mutations.append([layer, "INSERT", gid, None, "(new row)", None,
                                  "; ".join(f"{k}={v}" for k, v in ins.items() if v is not None)])
                coverage.append([layer, gid, r.get("Projectnaam"), "NEW_IN_EXCEL", geom_note])
                # stash full dict for apply phase
                mutations[-1].append({"values": ins, "x": x, "y": y})

        for gid, fid in gpkg_ids.items():
            if gid not in excel_ids:
                name_col = "zv_naam" if layer == "zov" else "wl_naam_p"
                coverage.append([layer, gid, gpkg_rows[fid].get(name_col),
                                 "ONLY_IN_GPKG", "No monitor row - left untouched (special case?)"])

    return mutations, warnings, coverage


def print_summary(mutations, warnings, coverage):
    upd = [m for m in mutations if m[1] == "UPDATE"]
    ins = [m for m in mutations if m[1] == "INSERT"]

    print("\n" + "=" * 70)
    print("PLAN SUMMARY")
    print("=" * 70)
    print(f"Cell updates: {len(upd)}   New projects to insert: {len(ins)}   Warnings: {len(warnings)}")

    if upd:
        counts = {}
        for m in upd:
            counts[(m[0], m[4])] = counts.get((m[0], m[4]), 0) + 1
        print("\nUpdates per column:")
        for (layer, col), n in sorted(counts.items()):
            print(f"  {layer:4s} {col:14s} {n:4d}")

    if ins:
        print("\nNew projects:")
        for m in ins:
            print(f"  {m[0]:4s} {m[2]}")

    cats = {}
    for c in coverage:
        cats.setdefault(c[3], []).append(c)
    print("\nCoverage:")
    for cat, rows in sorted(cats.items()):
        print(f"  {cat}: {len(rows)}")
        for r in rows:
            if cat in ("ONLY_IN_GPKG", "STOPPED_STILL_IN_GPKG", "NEW_IN_EXCEL"):
                print(f"    {r[0]:4s} {str(r[1]):12s} {str(r[2])[:45]:45s} {r[4]}")

    if warnings:
        print(f"\nWarnings ({len(warnings)}):")
        for w in warnings:
            print(f"  {w[0]:4s} {str(w[1]):12s} {str(w[2])[:35]:35s} {w[3]}")


def write_plan_files(outdir, mutations, warnings, coverage):
    with open(outdir / "plan_mutations.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["layer", "action", "project_id", "fid", "column", "old", "new"])
        for m in mutations:
            w.writerow(m[:7])
    with open(outdir / "plan_warnings.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["layer", "project_id", "project", "warning"])
        w.writerows(warnings)
    with open(outdir / "plan_coverage.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["layer", "project_id", "project", "category", "note"])
        w.writerows(coverage)
    print(f"\nPlan written to {outdir}/plan_mutations.csv, plan_warnings.csv, plan_coverage.csv")


# --------------------------------------------------------------------------
# Apply
# --------------------------------------------------------------------------

def apply_plan(gpkg_path, mutations, output_dir):
    # Never touch the input file: copy it into output/ and mutate the copy.
    output_dir.mkdir(parents=True, exist_ok=True)
    dest = output_dir / (gpkg_path.stem + "_updated_"
                         + datetime.now().strftime("%Y%m%d_%H%M%S") + gpkg_path.suffix)
    shutil.copy2(gpkg_path, dest)
    print(f"\nWriting updated gpkg to: {dest}")

    con = sqlite3.connect(dest)
    register_st_functions(con)
    cur = con.cursor()
    n_upd = n_ins = 0
    try:
        for m in mutations:
            layer, action = m[0], m[1]
            table = LAYERS[layer]["table"]
            if action == "UPDATE":
                _, _, gid, fid, col, _, new = m[:7]
                cur.execute(f'UPDATE "{table}" SET "{col}" = ? WHERE fid = ?', (new, fid))
                n_upd += 1
            elif action == "INSERT":
                payload = m[7]
                vals = payload["values"]
                x, y = payload["x"], payload["y"]
                geom = square_multipolygon_blob(x, y) if valid_rd(x, y) else None
                cols = ["geom"] + list(vals.keys())
                params = [geom] + list(vals.values())
                q = f'INSERT INTO "{table}" ({", ".join(chr(34)+c+chr(34) for c in cols)}) VALUES ({", ".join("?"*len(cols))})'
                cur.execute(q, params)
                n_ins += 1
        cur.execute("UPDATE gpkg_contents SET last_change = ? WHERE table_name IN (?, ?)",
                    (datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
                     LAYERS["zov"]["table"], LAYERS["wol"]["table"]))
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()
    print(f"Applied: {n_upd} cell updates, {n_ins} inserts. "
          f"Original input left untouched at {gpkg_path}.")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--gpkg", required=True, type=Path,
                    help=f"gpkg filename in {INPUT_DIR}/ (or a full path)")
    ap.add_argument("--excel", required=True, type=Path,
                    help=f"excel filename in {INPUT_DIR}/ (or a full path)")
    ap.add_argument("--sheet", default=None, help="Monitor sheet name (default: first sheet starting with 'Monitor')")
    ap.add_argument("--input-dir", type=Path, default=INPUT_DIR, help="folder to read files from")
    ap.add_argument("--output-dir", type=Path, default=OUTPUT_DIR, help="folder to write plan CSVs and updated gpkg to")
    ap.add_argument("--apply", action="store_true", help="Execute the plan (default is dry-run)")
    args = ap.parse_args()

    # Resolve filenames against the input folder, unless a full/existing path was given.
    gpkg_path = args.gpkg if args.gpkg.exists() else args.input_dir / args.gpkg
    excel_path = args.excel if args.excel.exists() else args.input_dir / args.excel
    for p, kind in ((gpkg_path, "gpkg"), (excel_path, "excel")):
        if not p.exists():
            sys.exit(f"{kind} file not found: {p}")
    args.output_dir.mkdir(parents=True, exist_ok=True)

    df, _ = load_monitor(excel_path, args.sheet)
    # normalize the 'Productie ' header (trailing space) once
    df.columns = ["Productie" if str(c).strip() == "Productie" else c for c in df.columns]

    con = sqlite3.connect(gpkg_path)
    mutations, warnings, coverage = build_plan(con, df)
    con.close()

    print_summary(mutations, warnings, coverage)
    write_plan_files(args.output_dir, mutations, warnings, coverage)

    if args.apply:
        confirm = input("\nType 'yes' to apply these mutations to the geopackage: ")
        if confirm.strip().lower() == "yes":
            apply_plan(gpkg_path, mutations, args.output_dir)
        else:
            print("Aborted - nothing changed.")
    else:
        print("\nDry-run only. Re-run with --apply to execute.")


if __name__ == "__main__":
    main()
