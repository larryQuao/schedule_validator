# Schedule Validator — Project Documentation

**Version 0.1.0 · 2026-09-27 · Java 21 + JavaFX desktop application**

Validates pension contribution report schedules against a reference workbook, and converts
non-Excel uploads (PDF, images) into the reference `.xlsx` layout.

- Repository: `github.com/larryQuao/schedule_validator` (branch `main`)
- Local project: `C:\Users\Administrator\Documents\Schedule_Validator`
- Quick-start / build / install commands live in [README.md](README.md); this document
  explains *what the system is, how it works, and why*.

---

## 1. Business context

The reference sample (`Samples/ANNE MARIE CONTRIBUTION REPORT.xlsx`, kept out of version
control — see §11) is an employer pension contribution workbook in the Ghana SSNIT style:

- **~106 monthly schedule sheets** (AUGUST 2017 → AUG 2026), one per contribution month.
- Each monthly sheet: a header block (SSNIT registration number, employer name, scheme
  name), a column-header row (**Member Code · SS No · Surname · Firstname · Other Names ·
  Basic Salary · 5% Contribution**, older years add *Description* and *Date of Birth*),
  member rows, and a printed totals row.
- **Auxiliary sheets**: `DETAILS` (member × month contribution matrix with a TOTAL column),
  `SUMMARY` (per-member totals), several `SUSPENSE …` sheets (unallocated/excess payments),
  `SURCHARGE`, and unnamed scratch sheets.
- Layout drifts across years: column positions shift, some rows lack Member Codes, a few
  sheets carry a second side-by-side table in columns M–S, and totals rows are sometimes
  followed by appended member rows.

The application ingests such reports — as `.xlsx`, PDF, or images — normalizes them into a
single domain model, applies validation rules, compares an upload against a reference, and
re-emits non-xlsx uploads as `.xlsx` in the reference layout.

---

## 2. Tech stack and rationale

| Concern | Choice | Rationale |
|---|---|---|
| Language / runtime | Java 21 LTS (Temurin) | Records + pattern matching fit the per-format adapter design; long support horizon |
| UI | JavaFX 21 | Desktop-native, `TableView` for findings, CSS theming, no browser involved |
| Excel I/O | Apache POI 5.3 | Reads `.xlsx` incl. merged cells/formulas; writes the converted layout faithfully |
| PDF tables | Apache PDFBox 3 (text positions) | Positional word extraction → line clustering → column mapping; avoids tabula-java's PDFBox 1.8 dependency clash |
| Image OCR | Tess4J 5.16 (Tesseract, offline) | Sensitive member data must not leave the machine; word bounding boxes feed the same table reconstructor as PDF |
| Build | Maven 3.9 | Conventional; `exec`/`javafx` plugins used for runs |
| Installer | `jpackage` + WiX Toolset 3.14 | MSI bundles a private JRE — zero prerequisites on target PCs |
| Tests | JUnit 5 | Integration tests run against the real sample workbook |

---

## 3. Architecture

One normalized domain model; every format funnels into it; validation and conversion only
ever see that model.

```
┌──────────────────────────────────────────────────────────────┐
│  ValidatorApp (JavaFX)            HeadlessRunner (CLI)       │
└───────────────┬──────────────────────────────┬───────────────┘
                └──────────┬───────────────────┘
                    Pipeline.run(reference, upload, outDir)
                           │
     parse ─► validate ─► compare ─► export
                           │
   ┌───────────────────────┼─────────────────────────────┐
   │ XlsxScheduleParser    │ PdfScheduleParser           │ ImageScheduleParser
   │ (Apache POI)          │ (PDFBox text positions)     │ (Tess4J OCR)
   └───────────┬───────────┴──────────────┬──────────────┘
               └──────────► ContributionReport ◄─────────┘
                                  │
              ValidationEngine (intra-report rules)
              ReportComparator (reference vs upload)
              XlsxExporter (converted schedule + findings workbook)
```

Flow for the primary use case: the reference report (`.xlsx`) and the uploaded file
(`.xlsx`/`.pdf`/image) each go through their parser → both become `ContributionReport`s →
`ValidationEngine` runs per-report rules on the upload → `ReportComparator` diffs
reference vs upload month-by-month → findings are written to
`Documents\Schedule Validator\out\validation_report.xlsx`, and non-xlsx uploads are
additionally re-emitted as `converted_schedule.xlsx`.

### Module map

| Package / file | Responsibility |
|---|---|
| `model/ContributionEntry` | One member row from any source (record): seq, member code, SS no, name parts, salary, contribution, provenance row number |
| `model/ScheduleSheet` | One monthly table: period, header row, entries, printed totals, post-total rows, in-sheet month label |
| `model/ContributionReport` | Whole parsed file: header info, monthly sheets, summary entries, suspense records, parse notes |
| `model/ValidationIssue` | One finding: severity (ERROR/WARNING/INFO), source, sheet, location, rule, message |
| `parse/XlsxScheduleParser` | Core workbook parser (see §4) |
| `parse/PdfScheduleParser` | Per-page word collection from PDFBox `TextPosition`s, y-clustering, table reconstruction |
| `parse/ImageScheduleParser` | Tesseract word boxes (page seg-mode 6), y-clustering, table reconstruction; OCR digit fixes (`O`→`0`) |
| `parse/TableExtractor` | Shared table reconstruction: header line found by keyword voting (SS/SURNAME/SALARY/CONTRIBUTION/MEMBER), header words merged into column tokens, data words assigned by x-proximity |
| `validate/ValidationEngine` | Intra-report rule set (§5) |
| `compare/ReportComparator` | Reference vs upload, per period (§6) |
| `convert/XlsxExporter` | Converted schedule in reference layout + findings workbook (§7) |
| `Pipeline` | Format dispatch by extension, orchestration, output paths (`tessdata` resolution, writable output dir) |
| `Main` | Launcher: no args → JavaFX GUI; args → `HeadlessRunner` |
| `ValidatorApp` | Desktop UI: file pickers, findings table (severity-coloured, filterable tooltips), progress, export buttons |

---

## 4. The XLSX parser in detail

The reference layout drifts between years, so columns are **mapped by header text, never
by fixed position**. Header cell text is normalized (uppercase, non-alphanumerics
stripped) and matched against a role dictionary:

`MEMBERCODE · SSNO/SSNITNO · SURNAME · FIRSTNAME · OTHERNAMES · BASICSALARY ·
CONTRIBUTION (5%…) · TOTAL · DESCRIPTION · DATEOFBIRTH · S/N`

Key behaviours:

- **Header detection** — scans the first 15 rows for a row containing an SS-number column
  plus a salary or contribution column. All monthly/summary/detail sheets put headers on
  row 3; the scan tolerates variation.
- **Side-table exclusion** — older sheets carry a second table in columns M–S. After the
  first mapped column, a run of more than 2 unmapped columns ends the table; mappings
  beyond it are ignored. Duplicate roles to the right never override earlier ones.
- **Unlabeled columns** — the 2026 layout has an unlabeled sequence column (A) and an
  unlabeled full-name column (G). Sequence: the leftmost column left of the first mapped
  column whose first data value is 1 (and keeps counting). Full name: an unmapped text
  column sitting between Other Names and Basic Salary.
- **Row classification** — a row is a *member row* when it has identity (SS no / surname /
  member code / full name) **and** an amount; a *totals row* when it has no identity and
  numeric salary+contribution (first such row only); blank runs of >12 rows end scanning.
- **Post-totals entries** — scanning continues after the totals row; member rows found
  below it are captured and flagged (this is how AUG 2026's five appended members are
  detected).
- **Report header block** — `NAME OF EMPLOYER:`, `NAME OF SCHEME:`, `SSNIT REGISTRATION`
  labels are matched in rows 1–5; the value is the next non-empty cell to the right.
- **Period parsing** — sheet names like `AUGUST 2017`, `AUG 2026`, `OCTOBER TO DECEMBER
  2019`, `MAR-20 TO JUL-20`, `DEC-20 & FEB-21`, `jan-jun 13 and penalty` are parsed to
  `yyyy-MM` start/end. Sheets are ordered by period. Unparseable names fall back to the
  raw name as key.
- **Special sheets** — names containing `SUSPENSE` are parsed as payer/amount/description
  records; `DETAILS` stores per-member rows (contribution = TOTAL column); `SUMMARY`
  stores per-member totals; anything without a detectable schedule header is recorded as
  a parse note (never a crash).

### PDF / image pipelines

Both feed `TableExtractor`. PDF: `PDFTextStripper` collects words with x/y per page;
words on a baseline merge into tokens; tokens cluster into lines. Image: Tess4J returns
word bounding boxes; clustering merges lines by y-center (±12 px). A header line needs ≥3
keyword hits. If no header is found, the page/line count is reported as a parse note —
scanned PDFs with no text layer are surfaced that way and should be routed through the
image pipeline (screenshot the page → image upload).

---

## 5. Validation rule catalog (per report)

Amounts compare with a one-cent epsilon (`0.011`) to absorb binary float representation —
otherwise matching is exact.

| Rule | Severity | Meaning |
|---|---|---|
| `TOTAL_SALARY_MISMATCH` | ERROR | Printed salary total ≠ computed sum of entries |
| `TOTAL_CONTRIBUTION_MISMATCH` | ERROR | Printed 5%-contribution total ≠ computed sum |
| `ENTRIES_BELOW_TOTALS` | ERROR | Member rows appear below the totals row, excluded from it (message includes the unaccounted amount) |
| `RATE_MISMATCH` | ERROR | Contribution ≠ 5% of that row's basic salary (message shows the effective rate) |
| `MISSING_AMOUNT` | WARNING | Row lacks salary or contribution |
| `NO_TOTALS_ROW` | WARNING | No recognizable totals row |
| `SEQ_GAP` / `SEQ_DUPLICATE` | WARNING | S/N skips or repeats |
| `MISSING_MEMBER_CODE` | WARNING | Rows without Member Code (count per sheet) |
| `MISSING_SS_NO` / `SS_FORMAT` | WARNING | No SS number, or one matching no known format (modern `[A-Z]\d{12}`, legacy `[A-Z0-9]{10,13}`) |
| `DUPLICATE_SS_IN_SHEET` | ERROR | Same SS number twice in one sheet |
| `MONTH_LABEL_MISMATCH` | WARNING | Sheet name month ≠ in-sheet `MONTH:` label |
| `NAME_WHITESPACE` | INFO | Leading/trailing/double spaces in name fields |
| `MEMBER_DISAPPEARED` | INFO | Member present in month N, absent in month N+1 |
| `SUMMARY_MISMATCH` (+ `SUMMARY_CROSSCHECK`) | WARNING | SUMMARY-sheet total for a member ≠ sum of that member's monthly entries |
| `PARSE_NOTE` | WARNING | Parser could not interpret part of the file |

### Comparison rules (reference vs upload)

Entries join on normalized SS number (uppercase, spaces stripped), falling back to
normalized full name. Months align on parsed period (`yyyy-MM`).

| Rule | Severity | Meaning |
|---|---|---|
| `SHEET_MISSING_IN_UPLOAD` | ERROR | Reference month absent from upload |
| `SHEET_NEW_IN_UPLOAD` | WARNING | Upload month not in reference |
| `ROW_MISSING_IN_UPLOAD` | ERROR | Member in reference, absent from upload month |
| `ROW_EXTRA_IN_UPLOAD` | WARNING | Member in upload month, not in reference |
| `SALARY_MISMATCH` / `CONTRIBUTION_MISMATCH` | ERROR | Same member, different amount (both values in message) |
| `NAME_MISMATCH` | WARNING | Normalized names differ |
| `STATED_TOTAL_DIFFERS` | WARNING | Printed month totals differ between the two files |

---

## 6. Outputs

**Sheet-merge mode (new-month workflow)** — when the user picks a specific sheet of the
reference report plus an upload, `Pipeline.runSheetMerge` + `SheetMerger` produce a *copy*
of the contribution report with a **new sheet for the uploaded month**, built from the
selected sheet as template:

- header block, column headers, per-column styles, merged regions, column widths and
  freeze panes are copied from the template;
- per-cell data-row formulas (`G=CONCATENATE(D," ",E," ",F)`, `I=0.05*H`, `J=IF(C=M,…`)
  are carried into the new rows with relative references shifted;
- members found in both files keep the template's order and take the upload's values;
- members only in the upload are appended and filled **green** (LIGHT_GREEN);
- members of the template missing from the upload are written **below the totals row**,
  filled **red** (ROSE) with `REMOVED` in the first column, and excluded from the totals;
- the totals row becomes a live `=SUM()` over kept + new rows only;
- output: `<reference name> - updated.xlsx` beside the other reports (the original file
  is never modified); the merge summary appears as an INFO finding.

**`validation_report.xlsx`** — two sheets: *Summary* (generation time, input files,
error/warning/info counts) and *Findings* (one row per issue: severity-coloured, with
source file kind, sheet, location, rule, message; auto-filter and frozen header).

**`converted_schedule.xlsx`** (pdf/image uploads) — the upload re-created in the reference
layout: `B1:C1` merge, row-2 header block (SSNIT reg / employer / scheme), row-3 column
headers, member rows from row 4, computed totals row, sample column widths, frozen panes.

All are written to `%USERPROFILE%\Documents\Schedule Validator\out` (the install
directory is not writable).

---

## 7. Verification performed so far

- **Build**: `mvn test` green — 4/4 integration tests against the real sample.
- **Parser**: all ~106 monthly sheets parsed; AUG 2026 yields its 38 entries and the
  printed totals; period parsing verified across name variants.
- **Rule engine on the sample**: 301 errors, 326 warnings, 236 info — all genuine sample
  defects (e.g. AUG 2026 totals exclude 5 appended members: printed 1,290.00 vs computed
  1,460.00; the sheet named AUGUST 2017 is labelled `MONTH:SEPTEMBER`; rows paid at 5.5%
  instead of 5%; SS `C017803280209` duplicated within single months).
- **Comparator**: self-comparison (reference = upload = sample) produces **zero**
  comparison findings, as expected.
- **GUI**: launches from the classpath (the installed-app mode) and via `mvn javafx:run`.
- **Installer**: MSI built with `jpackage`/WiX 3.14, silently installed to
  `C:\Program Files\Schedule Validator`, and the installed app verified running with the
  OCR data in place (`app\tessdata\eng.traineddata`).

---

## 8. Packaging and installation

`mvn package` produces the app jar and copies runtime dependencies to `target\libs`.
The staged `installer\input` (app jar + `libs\` + `tessdata\`) is fed to `jpackage`
(`--type msi --win-menu --win-shortcut --win-dir-chooser`), which requires **WiX Toolset
3.14** (`winget install WiXToolset.WiXToolset`). Result:
`installer\out\Schedule Validator-0.1.0.msi` (~108 MB) — bundles a private JRE, JavaFX,
all libraries, and OCR data; installs Start-Menu and desktop shortcuts; no Java needed on
target machines. The installed app resolves `tessdata` next to its jars and writes
outputs to the user's Documents folder.

---

## 9. Development environment (as set up on this workstation)

| Tool | Location |
|---|---|
| JDK 21 (Temurin 21.0.12.1+1) | `C:\Users\Administrator\tools\jdk-21.0.12.1+1` |
| Maven 3.9.9 | `C:\Users\Administrator\tools\apache-maven-3.9.9` |
| WiX Toolset 3.14 | machine-wide (jpackage dependency) |
| Git for Windows 2.55, GitHub CLI | machine-wide; git identity is repo-local (`larryQuao`), push auth via SSH key `%USERPROFILE%\.ssh\id_ed25519` |

Neither JDK nor Maven is on the system PATH — commands in this doc assume `JAVA_HOME`
set and both `bin` dirs prepended (see README). Note: in `cmd`, `%VAR%` set on the same
command line expands before `set` takes effect; use a literal path or separate lines.

---

## 10. Known limitations

- **Scanned PDFs** have no text layer — PDFBox finds no words. Export a page image and
  upload it as PNG/JPG instead.
- **OCR accuracy** depends on scan quality (clean horizontal scans work best); every
  converted row keeps its source line number for manual checking.
- **Side tables** (columns M–S on older sheets) are deliberately ignored.
- **DETAILS** stores only per-member totals, not the monthly matrix columns.
- **SURCHARGE** and unnamed scratch sheets (`Sheet5`) are reported as parse notes — their
  semantics are unconfirmed.
- Period parsing is best-effort; exotic sheet names fall back to the raw name.
- Exact-match only — no tolerance/fuzzy mode yet (§11).
- The 2017-era mixed 5%/5.5% rates may be *historically correct*; treat `RATE_MISMATCH`
  findings in that era as review prompts, not necessarily errors.

## 11. Roadmap

1. Tune PDF/image conversion against real uploads (page headers, multi-page months).
2. Confirm SURCHARGE / scratch-sheet semantics and parse them.
3. Optional tolerance mode (rule tiers instead of exact match).
4. Historical-rate configuration (5.5% era) so old rows aren't flagged.
5. `jpackage` build wired into Maven; signed installer.
6. Data handling: the sample workbook contains real member personal data and is excluded
   from version control by `.gitignore` — keep it that way, or replace with an
   anonymized fixture if the repo must be self-contained.
