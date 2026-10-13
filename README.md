# ENABLE EMR Export Loader

Antenatal care visits recorded in the EMR need to reach the DHIS2
ANC tracker, where they drive an SMS programme of lifestyle messages and
appointment reminders. Until the facilities push their data directly, their
exports travel as CSV files. This tool converts each
facility's export into the format the ENABLE mediator accepts and sends the
records to the mediator over HTTPS. The mediator, which runs alongside the
DHIS2 server, validates every record and imports it into the DHIS2 tracker program.

```
facility export (.xlsx) ──► incoming/ ──loader.py──► mediator ──► DHIS2 ANC tracker
```

The loader cleans up what can be cleaned mechanically (like normalizing the GA entries to a single format) and rejects what it
cannot. Phone numbers are normalised to the international format the SMS
programme requires, free-text gestational ages ("22", "22w+4d", "22 wks")
become whole weeks, and duplicate visit entries keep their newest version. A
record that cannot be imported truthfully, e.g. by not having a usable visit number, no way to date the pregnancy, or no valid phone, is left out and
listed in the conversion report, by medical record number and reason, so the
facility can fix the source data. Imports are idempotent: re-sending an
export updates the records it already created, it never duplicates them, so
running the same file twice is always safe.

These files contain patient data. Keep this folder on an encrypted computer,
out of shared drives and email, and delete dumps from `processed/` once an
import is confirmed. The `.gitignore` is set up so that neither patient data
nor the API key can reach git.

## Setting up

The tool needs Python 3.9 or newer and keeps its libraries in a private
`.venv` folder inside this directory. Nothing is installed system-wide and
no administrator rights are needed.

On a Mac, open Terminal in this folder and check `python3 --version`. If that
fails or shows a version below 3.9, install Python with
[Homebrew](https://brew.sh):

```sh
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
brew install python
```

then close and reopen Terminal. Create the library folder and install:

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

On Windows, install Python from <https://www.python.org/downloads/> (tick
**"Add python.exe to PATH"** in the installer), then in PowerShell in this
folder:

```powershell
py -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

Finally, on either system, copy `config.env.example` to `config.env` and fill
in the mediator's URL and the PulseTech API key. Both come from the
integration team, and the key should be treated like a password. Check
`facilities.csv`: one line per facility, where `match` is a word that appears
in that facility's file names and `facility_code` is its DHIS2 code, provided
when the facility is onboarded. Adding a facility means adding a line.

```csv
match,facility_code,facility_name
akaki,1029107,Akaki Health Center
```

## Converting and sending

Put the facility's `.xlsx` file(s) in `incoming/` and run, from this folder:

```sh
.venv/bin/python3 loader.py        # Mac
.venv\Scripts\python loader.py     # Windows
```

(Calling the `.venv` Python directly means there is no environment to
remember to activate.) For each file the loader prints how many women were
converted and pushed and where the reports are. A file that went through
completely moves to `processed/` with a date stamp; a file with a problem
stays in `incoming/` so that, once the cause is fixed, the same command picks
it up again. Add `--dry-run` to convert and write reports without sending
anything.

## The reports

Each run writes a folder `out/<facility code>-<timestamp>/`:

| File | What it holds |
|---|---|
| `...-report.md` | The conversion report: what was imported, and every record left out or adjusted, grouped by consequence. Women not imported (with what to fix), visit rows left out, things imported but worth checking. This should be reported back to facilities to help identify possible data issues. |
| `push-report.txt` | The mediator's answer per batch. HTTP status, timing, per-record outcome counts, and any rejection quoted with its reason. |
| `response-*.json` | The mediator's full answers, kept for reference. |

## When something goes wrong

| What you see | What it means | What to do |
|---|---|---|
| `SKIPPED: no line in facilities.csv matches` | The loader cannot tell which facility the file is from | Add a line to `facilities.csv` (ask the team for the facility code) |
| `HTTP 403 — the API key was refused` | Wrong or missing key | Check `MEDIATOR_API_KEY_PULSETECH` in `config.env` with the team |
| `Mediator unreachable` | No connection to the server | Check the internet connection and `MEDIATOR_URL`; if both look right, tell the team |
| `HTTP 502` | The mediator is up but DHIS2 is not responding | Run the same command again later. Re-pushing the same data is safe |
| `FAILED to convert` | The Excel file does not have the expected sheets | The export is incomplete or from a different system. Go back to the facility |
| Many records rejected for the same reason | A data-quality problem at the facility | Send the conversion report to the facility; re-send the corrected export later |
| Certificate error when pushing | The server's certificate is not trusted | Ask the team; only for a test server should `MEDIATOR_TLS_VERIFY=false` be set |

## Trying it out safely

`test-data/Test_Facility_Data.xlsx` is a synthetic export, including eleven dummy patient records (MRNs 900001–900011, fake phone numbers) carrying one of every known
data problem. Copy it into `incoming/` and run the loader to see the whole
flow, rejections included. 8 of the 11 import, and the conversion report
shows every failure category with its reason. The file is matched by the
`test_facility` line in `facilities.csv`; point its code at a test instance's
org unit, never at production. The imported test women are easy to find and
delete afterwards: their MRNs start with 900.

## Support

For any questions or troubleshooting, please contact me at johan@devotta.no
