#!/usr/bin/env python3
"""Converts the PulseTech facility exports in incoming/ and pushes the records
to the mediator, which validates each one and imports it into the DHIS2 ANC
tracker.

  python3 loader.py             # process everything in incoming/
  python3 loader.py --dry-run   # convert and report only, push nothing

Each export is matched to its facility by file name (facilities.csv), converted
into batch files under out/<facility>-<timestamp>/ together with the conversion
and push reports, and moved to processed/ once every batch is through. An
export that fails stays in incoming/ so the next run retries it; imports are
idempotent, so re-pushing updates records rather than duplicating them.

Configuration lives in config.env (copy config.env.example to start).
"""

import argparse
import csv
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

import requests

from convert import convert_dump

HERE = Path(__file__).resolve().parent
PUSH_TIMEOUT_SECONDS = 900  # a large batch takes a while; the mediator answers when done


def load_config(path):
    if not path.exists():
        sys.exit(f"Missing {path.name} — copy config.env.example to config.env and fill it in.")
    cfg = {}
    for line in path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, _, value = line.partition("=")
            cfg[key.strip()] = value.strip()
    for required in ("MEDIATOR_URL", "MEDIATOR_API_KEY_PULSETECH"):
        if not cfg.get(required):
            sys.exit(f"{required} is not set in config.env.")
    return cfg


def load_facilities(path):
    if not path.exists():
        sys.exit(f"Missing {path.name} — one line per facility, see the README.")
    with path.open(newline="") as fh:
        rows = [r for r in csv.DictReader(fh) if r.get("match", "").strip()]
    if not rows:
        sys.exit(f"{path.name} has no facility lines.")
    return rows


def match_facility(filename, facilities):
    """First facilities.csv row whose 'match' text appears in the file name."""
    name = filename.lower()
    for row in facilities:
        if row["match"].strip().lower() in name:
            return row
    return None


SUCCESS_STATUSES = {"OK", "SUCCESS", "IMPORTED", "CREATED", "UPDATED"}


def summarize_response(body_text):
    """-> (counter of per-record statuses, list of rejected-record notes)."""
    counts, rejected = {}, []
    try:
        body = json.loads(body_text)
    except (ValueError, TypeError):
        return counts, rejected
    results = body.get("results") if isinstance(body, dict) else None
    for r in results if isinstance(results, list) else []:
        status = str(r.get("status", "?")).upper()
        counts[status] = counts.get(status, 0) + 1
        if status not in SUCCESS_STATUSES:
            rejected.append(json.dumps(r, ensure_ascii=False)[:500])
    return counts, rejected


def push_batches(batch_files, cfg, run_dir):
    """Push every batch file. Returns (ok, lines-for-the-push-report)."""
    url = cfg["MEDIATOR_URL"].rstrip("/") + "/api/pulsetech/anc-records"
    verify = cfg.get("MEDIATOR_TLS_VERIFY", "true").strip().lower() != "false"
    if not verify:
        requests.packages.urllib3.disable_warnings()
    headers = {"X-API-KEY": cfg["MEDIATOR_API_KEY_PULSETECH"],
               "Content-Type": "application/json"}

    lines, all_ok = [f"Pushed to {url}", ""], True
    for batch in batch_files:
        started = datetime.now()
        try:
            response = requests.post(url, data=batch.read_bytes(), headers=headers,
                                     verify=verify, timeout=PUSH_TIMEOUT_SECONDS)
        except requests.RequestException as exc:
            lines.append(f"FAILED  {batch.name}: {exc}")
            lines.append("        Mediator unreachable — stopping; the file stays in incoming/.")
            return False, lines
        seconds = (datetime.now() - started).total_seconds()
        (run_dir / f"response-{batch.stem}.json").write_text(response.text)

        # The mediator's error contract (see the main repository's README):
        # 200 all imported · 207 mixed, per-record statuses in results[] ·
        # 400 invalid record · 401/403 key problem · 409 DHIS2 rejected the
        # import · 413/429 too large / too fast (safe to retry) · 502 DHIS2
        # unreachable. Imports are idempotent, so retrying is always safe.
        if response.status_code in (401, 403):
            lines.append(f"FAILED  {batch.name}: HTTP {response.status_code} — the API key was "
                         "refused. Check MEDIATOR_API_KEY_PULSETECH in config.env; stopping.")
            return False, lines
        if response.status_code == 502:
            lines.append(f"FAILED  {batch.name}: HTTP 502 — the mediator could not reach DHIS2. "
                         "Stopping; run again later (re-pushing is safe).")
            return False, lines
        counts, rejected = summarize_response(response.text)
        note = f" {counts}" if counts else ""
        lines.append(f"{response.status_code}  {batch.name}  ({seconds:.0f}s){note}")
        if response.status_code not in (200, 207):
            all_ok = False
            hint = {400: "the mediator rejected a record as invalid — see the response file",
                    409: "DHIS2 rejected the import; its report is in the response file",
                    413: "batch too large — safe to retry with a smaller --batch-size",
                    429: "too many requests — safe to retry"}.get(response.status_code)
            if hint:
                lines.append(f"        {hint}")
        lines.extend(f"        rejected: {r}" for r in rejected[:50])
    return all_ok, lines


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dry-run", action="store_true",
                    help="convert and write reports, but push nothing")
    ap.add_argument("--batch-size", type=int, default=200)
    args = ap.parse_args()

    cfg = load_config(HERE / "config.env")
    facilities = load_facilities(HERE / "facilities.csv")
    dumps = sorted((HERE / "incoming").glob("*.xlsx"))
    if not dumps:
        print("Nothing to do: no .xlsx files in incoming/.")
        return

    failures = 0
    for dump in dumps:
        print(f"\n=== {dump.name} ===")
        facility = match_facility(dump.name, facilities)
        if facility is None:
            print("  SKIPPED: no line in facilities.csv matches this file name. "
                  "Add one (match,facility_code,facility_name) and run again.")
            failures += 1
            continue

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        run_dir = HERE / "out" / f"{facility['facility_code']}-{stamp}"
        try:
            summary = convert_dump(str(dump), facility["facility_code"],
                                   facility.get("facility_name", ""),
                                   out_dir=str(run_dir), batch_size=args.batch_size)
        except Exception as exc:  # a malformed workbook must not stop the other files
            print(f"  FAILED to convert: {exc}")
            failures += 1
            continue
        print(f"  {summary['records']} of {summary['women']} women converted "
              f"-> {len(summary['batch_files'])} batch file(s)")
        print(f"  conversion report: {summary['report']}")
        if summary["records"] == 0:
            print("  SKIPPED push: nothing to send (see the report).")
            failures += 1
            continue

        if args.dry_run:
            print("  dry run: not pushing; file stays in incoming/.")
            continue

        ok, lines = push_batches(summary["batch_files"], cfg, run_dir)
        (run_dir / "push-report.txt").write_text("\n".join(lines) + "\n")
        print("  " + "\n  ".join(lines[-min(len(lines), 6):]))
        print(f"  push report: {run_dir / 'push-report.txt'}")
        if ok:
            destination = HERE / "processed" / f"{stamp}-{dump.name}"
            shutil.move(str(dump), destination)
            print(f"  done: moved to processed/{destination.name}")
        else:
            failures += 1
            print("  NOT moved: fix the problem above and run again (re-pushing is safe).")

    if failures:
        sys.exit(f"\nFinished with {failures} file(s) needing attention.")
    print("\nAll files processed.")


if __name__ == "__main__":
    main()
