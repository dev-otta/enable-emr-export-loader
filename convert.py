#!/usr/bin/env python3
"""PulseTech facility data dump (xlsx) -> mediator batch JSON.

Generic per facility: everything facility-specific is a CLI argument.

  python3 pulsetech_dump_to_json.py Akaki_HC_Data.xlsx \
      --facility-code 1029107 --facility-name "Akaki Health Center" \
      --out-dir out --batch-size 200

Reads the Patient, anc-visit, anc-followup and Obstetric_Traige tabs
(everything else is ignored). Writes:
  out/<facility>-anc-records-NNN.json   mediator batch files ({"records":[...]})
  out/<facility>-report.md              every row/woman that needed fixing or skipping

Decisions it makes (all counted in the report):
  phone     digits only; Ethiopian numbers are normalised to 2519XXXXXXXX /
            2517XXXXXXXX. Unparseable phone => record skipped (mandatory, SMS target).
  age       the dump has age, not date of birth: pnt_age is sent as-is
            (optional Age (years) attribute); no DOB is sent or estimated.
  ga        free text ("22", "22w+4d", "22+4", "22 wks") -> whole weeks (days ignored).
  visit_num "Visit N" -> N. "Visit >=8" -> 8, 9, ... in date order per woman.
            Yes/No/blank => row skipped. Duplicate N per woman => latest row kept.
  lmp/edd   OPTIONAL — never derived. Taken per woman from the first followup row
            with a plausible pair (EDD-LMP 200-320d), else Obstetric_Traige;
            omitted when neither has one.
  ga        MANDATORY per entry (contract rule): a row without a parseable GA is
            skipped — better not imported than imported undatable.
"""

import argparse
import json
import math
import re
import sys
from collections import defaultdict
from datetime import timedelta
from pathlib import Path

import pandas as pd

GA_MIN_WEEKS, GA_MAX_WEEKS = 4, 45
PLAUSIBLE_EDD_MINUS_LMP = (200, 320)  # days; 280 is the textbook value


def parse_ga(raw):
    """'22', '22w+4d', '22+4', '22 wks', '22wk' -> 22. None when unusable."""
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return None
    m = re.match(r"\s*(\d+)", str(raw))
    if not m:
        return None
    weeks = int(m.group(1))
    return weeks if GA_MIN_WEEKS <= weeks <= GA_MAX_WEEKS else None


def parse_visit_num(raw):
    """'Visit 3' -> (3, False); 'Visit >=8' -> (8, True); junk -> (None, False)."""
    s = str(raw).strip()
    m = re.fullmatch(r"Visit (\d+)", s)
    if m:
        return int(m.group(1)), False
    if re.fullmatch(r"Visit >=\s*(\d+)", s):
        return int(re.search(r"\d+", s).group()), True
    return None, False


def normalize_phone(raw, nationality):
    """Ethiopian numbers -> 251XXXXXXXXX. Returns None when unusable."""
    digits = re.sub(r"\D", "", str(raw))
    if not str(nationality).strip().lower().startswith("ethiopian"):
        return digits if len(digits) >= 7 else None
    if digits.startswith("251") and len(digits) == 12:
        return digits
    if digits.startswith("0") and len(digits) == 10:
        digits = digits[1:]
    if len(digits) == 9 and digits[0] in "97":
        return "251" + digits
    return None


def iso(d):
    return pd.Timestamp(d).strftime("%Y-%m-%d")


def convert_dump(dump, facility_code, facility_name="", out_dir="out", batch_size=200):
    """Convert one PulseTech xlsx dump to mediator batch files.

    Returns {"records", "women", "batch_files", "report"}."""
    xl = pd.ExcelFile(dump)
    pat = xl.parse("Patient")
    fup = xl.parse("anc-followup")
    try:
        triage = xl.parse("Obstetric_Traige")
    except Exception:
        triage = pd.DataFrame(columns=["MRN"])
    try:
        visit = xl.parse("anc-visit")
    except Exception:
        visit = pd.DataFrame(columns=["MRN", "exam_date"])

    report = defaultdict(list)

    # ---- patients ---------------------------------------------------------
    patients = {}
    for p in pat.itertuples(index=False):
        mrn = int(p.mrn)
        phone = normalize_phone(p.pnt_phone, p.pnt_nationality)
        if phone is None:
            report["invalid_phone (skipped)"].append(mrn)
            continue
        patients[mrn] = {
            "pnt_mrn": str(mrn),
            "pnt_gender": str(p.pnt_gender),
            "pnt_phone": phone,
        }
        if str(p.pnt_age_type).strip().lower() == "years":
            patients[mrn]["pnt_age"] = int(p.pnt_age)
        else:
            report["age_not_in_years (age omitted)"].append(mrn)
        if str(p.pnt_gender).strip().lower() != "female":
            report["gender_not_female (kept, verify)"].append(mrn)

    phones = defaultdict(list)
    for mrn, p in patients.items():
        phones[p["pnt_phone"]].append(mrn)
    for phone, mrns in phones.items():
        if len(mrns) > 1:
            report["phone_shared_by_multiple_mrns (kept, verify)"].append(
                f"{phone[:5]}****: MRNs {sorted(mrns)}")

    # ---- canonical LMP/EDD per woman -------------------------------------
    def plausible(lmp, edd):
        if pd.isna(lmp) or pd.isna(edd):
            return False
        return PLAUSIBLE_EDD_MINUS_LMP[0] <= (edd - lmp).days <= PLAUSIBLE_EDD_MINUS_LMP[1]

    triage_dates = {}
    if {"MRN", "centpreg_lmp_date", "centpreg_edd_date"} <= set(triage.columns):
        for t in triage.sort_values("Date").itertuples(index=False):
            lmp = pd.to_datetime(t.centpreg_lmp_date, errors="coerce")
            edd = pd.to_datetime(t.centpreg_edd_date, errors="coerce")
            if plausible(lmp, edd):
                triage_dates.setdefault(int(t.MRN), (lmp, edd))

    # ---- followups --------------------------------------------------------
    f = fup.dropna(subset=["MRN"]).copy()
    report["followup_missing_mrn (rows skipped)"] = [f"{len(fup) - len(f)} rows"]
    f["MRN"] = f["MRN"].astype(int)
    f = f.sort_values(["MRN", "date", "created_at"])

    records, skipped_women = [], defaultdict(list)
    for mrn, rows in f.groupby("MRN"):
        if mrn not in patients:
            skipped_women["no_usable_patient_row"].append(mrn)
            continue

        # visit numbers: parse, spread '>=8' sequentially, drop junk, dedupe
        entries, ge8_next, seen = [], 8, {}
        for r in rows.itertuples(index=False):
            num, is_ge8 = parse_visit_num(r.visit_num)
            if num is None:
                report["visit_num_unusable (row skipped)"].append(f"MRN {mrn}: {r.visit_num!r}")
                continue
            if is_ge8:
                num, ge8_next = max(ge8_next, num), max(ge8_next, num) + 1
            if num in seen:
                report["duplicate_visit_number (older row dropped)"].append(f"MRN {mrn}: Visit {num}")
            seen[num] = r  # later rows win (sorted by date, created_at)
        entries = [seen[k] for k in sorted(seen)]
        if not entries:
            skipped_women["no_usable_followup_rows"].append(mrn)
            continue

        # canonical pregnancy dates — real sources only, never derived
        lmp = edd = None
        for r in entries:
            if plausible(r.lmp, r.edd):
                lmp, edd = r.lmp, r.edd
                break
        if lmp is None and mrn in triage_dates:
            lmp, edd = triage_dates[mrn]
            report["lmp_edd_from_triage"].append(mrn)
        if lmp is None:
            report["no_lmp_edd (record sent without them)"].append(mrn)

        followups = []
        for num, r in sorted(seen.items()):
            ga = parse_ga(r.ga)
            if ga is None:
                report["ga_unusable (row skipped)"].append(f"MRN {mrn}: ga={r.ga!r}")
                continue
            entry = {
                "followup_date": iso(r.date),
                "ga": str(ga),
                "visit_number": f"Visit {num}",
            }
            if lmp is not None:
                entry["lmp_date"] = iso(lmp)
                entry["edd"] = iso(edd)
            followups.append(entry)
        if not followups:
            skipped_women["no_followup_row_with_usable_ga"].append(mrn)
            continue

        rec = dict(patients[mrn])
        rec["facility_dhis2_org_unit_code"] = facility_code
        if facility_name:
            rec["facility_name"] = facility_name
        v = visit[visit.MRN == mrn] if "MRN" in visit.columns else visit.iloc[0:0]
        rec["anc_visit"] = [{"visit_date": iso(d)} for d in sorted(v.exam_date)] if len(v) else []
        rec["anc_followup"] = followups
        records.append(rec)

    for k, v in skipped_women.items():
        report[f"woman_skipped: {k}"] = v

    # ---- write batches ----------------------------------------------------
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"\W+", "-", (facility_name or facility_code)).strip("-").lower()
    batch_files = []
    for i in range(0, len(records), batch_size):
        path = out / f"{stem}-anc-records-{len(batch_files) + 1:03d}.json"
        path.write_text(json.dumps({"records": records[i:i + batch_size]},
                                   indent=2, ensure_ascii=False) + "\n")
        batch_files.append(path)
    n_files = len(batch_files)

    # ---- report -----------------------------------------------------------
    # Grouped by consequence, most important first. Keys are mapped to readable
    # headings here; anything unmapped still prints at the end, so no finding
    # is ever silently dropped.
    LAYOUT = [
        ("Women NOT imported — to fix at the facility", [
            ("woman_skipped: no_usable_followup_rows",
             "All visit rows lack a visit number",
             "Every row for these women has something other than 'Visit N' in the "
             "visit number column. Fix the entries in PulseTech and re-export."),
            ("woman_skipped: no_followup_row_with_usable_ga",
             "No visit with a usable gestational age",
             "No row carries a readable GA, and there is no LMP/EDD to fall back "
             "on. Record the gestational age and re-export."),
            ("woman_skipped: no_usable_patient_row",
             "No usable patient register row",
             "Usually an unusable phone number — see the phone list below."),
            ("invalid_phone (skipped)",
             "Unusable phone number",
             "The SMS programme needs a valid number. Correct it and re-export."),
        ]),
        ("Visit rows left out (their women were still imported)", [
            ("visit_num_unusable (row skipped)", "Visit number not usable", None),
            ("ga_unusable (row skipped)", "Gestational age not usable", None),
            ("followup_missing_mrn (rows skipped)", "Visit row without an MRN", None),
            ("duplicate_visit_number (older row dropped)",
             "Same visit number entered twice (newest kept)", None),
        ]),
        ("Imported, but worth checking", [
            ("phone_shared_by_multiple_mrns (kept, verify)",
             "Phone number shared by several women",
             "These women will receive each other's SMS reminders."),
            ("gender_not_female (kept, verify)", "Gender is not female", None),
            ("age_not_in_years (age omitted)",
             "Age not recorded in years (left out of the record)", None),
        ]),
        ("For information", [
            ("lmp_edd_from_triage", "LMP/EDD taken from the triage sheet", None),
            ("no_lmp_edd (record sent without them)",
             "No LMP/EDD recorded — sent with the gestational age only", None),
        ]),
    ]

    def mrn_lines(values):
        values = sorted(str(v) for v in values)
        return [", ".join(values[i:i + 12]) for i in range(0, len(values), 12)]

    lines = [
        f"# {facility_name or facility_code} — conversion report",
        "",
        f"Of the {f.MRN.nunique()} women in this export, {len(records)} were "
        f"converted ({sum(len(r['anc_followup']) for r in records)} visits, "
        f"{n_files} batch file(s) of up to {batch_size} records). Everything "
        "left out or adjusted is listed below, most important first.",
    ]
    known = set()
    for section, entries in LAYOUT:
        known.update(key for key, _, _ in entries)
        picked = [(key, title, action) for key, title, action in entries if report.get(key)]
        if not picked:
            continue
        lines += ["", f"## {section}"]
        for key, title, action in picked:
            values = report[key]
            lines += ["", f"### {title} — {len(values)}"]
            if action:
                lines += [action]
            if all(isinstance(v, int) for v in values):
                lines += [""] + mrn_lines(values)
            else:
                lines += [f"- {v}" for v in values[:100]]
                if len(values) > 100:
                    lines.append(f"- ... and {len(values) - 100} more")
    for key in sorted(k for k, v in report.items() if k not in known and v):
        lines += ["", f"## {key} — {len(report[key])}"]
        lines += [f"- {v}" for v in report[key][:100]]

    report_path = out / f"{stem}-report.md"
    report_path.write_text("\n".join(lines) + "\n")

    return {
        "records": len(records),
        "women": int(f.MRN.nunique()),
        "batch_files": batch_files,
        "report": report_path,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("dump", help="PulseTech xlsx export")
    ap.add_argument("--facility-code", required=True, help="DHIS2 org unit CODE")
    ap.add_argument("--facility-name", default="")
    ap.add_argument("--out-dir", default="out")
    ap.add_argument("--batch-size", type=int, default=200, help="records per output file")
    args = ap.parse_args()
    s = convert_dump(args.dump, args.facility_code, args.facility_name, args.out_dir, args.batch_size)
    print(f"{s['records']} records -> {len(s['batch_files'])} file(s); report: {s['report']}",
          file=sys.stderr)


if __name__ == "__main__":
    main()
