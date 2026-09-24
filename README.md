# Schedule Validator

A Windows desktop application that validates pension **contribution report schedules**
against a reference workbook, and converts non-Excel uploads (PDF, images) into the
reference `.xlsx` layout.

Tuned against the reference sample `Samples/ANNE MARIE CONTRIBUTION REPORT.xlsx`
(SSNIT-style employer report: one sheet per month, Member Code / SS No / Surname /
Firstname / Other Names / Basic Salary / 5% Contribution columns, printed totals row).

## Stack

| Concern        | Choice                                      |
|----------------|---------------------------------------------|
| Language       | Java 21 (Temurin)                           |
| UI             | JavaFX 21                                   |
| Excel I/O      | Apache POI 5.3                              |
| PDF tables     | Apache PDFBox 3                             |
| Image OCR      | Tesseract via Tess4J (offline)              |
| Build          | Maven 3.9                                   |
| Tests          | JUnit 5 (runs against the real sample)      |

## What it does

1. **Parse** — `.xlsx` reference reports (all monthly sheets + DETAILS/SUMMARY/SUSPENSE),
   text-based PDFs (positional table reconstruction) and images/scans (OCR + column
   reconstruction). Columns are mapped by header text, so layout drift between years
   does not matter.
2. **Validate** (rule engine, exact-match by default):
   - printed totals vs computed totals (salary and 5% contribution)
   - member rows placed *below* the totals row (excluded from the printed total)
   - contribution ≠ 5% of basic salary (rate mismatches)
   - duplicate SS numbers within a sheet
   - sequence number gaps/duplicates, missing member codes/SS numbers
   - sheet name vs in-sheet `MONTH:` label disagreement
   - name whitespace irregularities, member disappearances month-over-month
   - SUMMARY sheet cross-check against the sum of monthly sheets
3. **Compare** — uploaded file vs reference, joined on SS number (name fallback),
   month-by-month: missing/extra members, salary/contribution/value differences.
4. **Convert** — PDF/image uploads are re-emitted as `.xlsx` in the reference layout
   (header block, labeled columns, computed totals row).
5. **Report** — `out/validation_report.xlsx`: Summary sheet + filterable Findings sheet
   (severity-coloured), one row per finding with sheet/row/rule/message.

## Running

### Windows installer (recommended)

Build once with:

```bat
mvn package -DskipTests                       :: jar + target\libs
:: stage installer\input (app jar, libs\, tessdata\), then:
jpackage --type msi --name "Schedule Validator" --app-version 0.1.0 ^
  --vendor "Schedule Validator Project" --input installer\input ^
  --main-jar schedule-validator-0.1.0.jar --main-class com.schedulevalidator.Main ^
  --dest installer\out --win-menu --win-shortcut --win-dir-chooser
```

This produces `installer\out\Schedule Validator-0.1.0.msi` (~108 MB: bundles its own
JRE, JavaFX, all libraries and OCR data — no Java needed on target machines).
Requires WiX Toolset 3.14 (`winget install WiXToolset.WiXToolset`) for MSI output.
Install: double-click the MSI (or `msiexec /i <msi> /qn` silently). It adds Start Menu
and desktop shortcuts, installing to `C:\Program Files\Schedule Validator`; validation
output is written to `%USERPROFILE%\Documents\Schedule Validator\out`.

### Development

Prerequisites: JDK 21 and Maven 3.9 on PATH
(`C:\Users\Administrator\tools\jdk-21.0.12.1+1` and `...\apache-maven-3.9.9` currently
live outside PATH — set `JAVA_HOME` and prepend both `bin` dirs, or use full paths).

```bat
:: desktop app (window with file pickers, findings table, export buttons)
mvn javafx:run

:: headless
java -cp target\classes;<deps> com.schedulevalidator.HeadlessRunner -r "Samples\ANNE MARIE CONTRIBUTION REPORT.xlsx" -u upload.pdf -o out

:: tests (include the sample-based integration tests)
mvn test
```

Headless options: `-r/--reference <file>` (existing report), `-u/--uploaded <file>`
(`.xlsx` `.pdf` `png/jpg/tiff/bmp`), `-o/--out <dir>` (default `out`).

OCR note: `tessdata\eng.traineddata` ships in the repo; the OCR path also needs the
Microsoft Visual C++ 2015–2022 x64 redistributable (usual case on Windows 10/11).
A log4j "no logging implementation" line at startup is harmless POI noise.

## Project layout

```
src/main/java/com/schedulevalidator/
  model/      ContributionEntry, ScheduleSheet, ContributionReport, ValidationIssue
  parse/      XlsxScheduleParser, PdfScheduleParser, ImageScheduleParser, TableExtractor
  validate/   ValidationEngine        (intra-report rules)
  compare/    ReportComparator        (reference vs upload, month by month)
  convert/    XlsxExporter            (converted schedule + findings workbook)
  Pipeline, HeadlessRunner, Main, ValidatorApp
src/main/resources/css/app.css
tessdata/eng.traineddata
Samples/ANNE MARIE CONTRIBUTION REPORT.xlsx
```

## Current results on the sample

The engine flags the sample's genuine defects, e.g.: AUG 2026 holds 5 members below
its totals row (printed 1290.00 vs computed 1460.00); the sheet named AUGUST 2017 is
labelled `MONTH:SEPTEMBER` inside; several 2017–2019 rows were paid at 5.5% instead of
5%; SS C017803280209 appears twice in single months. Self-comparison (reference =
upload = sample) yields zero comparison findings, as expected.

## Roadmap

- DETAILed PDF/image tuning against real uploads (page headers, multi-page months)
- SURCHARGE / "Sheet5" layouts once their semantics are confirmed
- optional tolerance mode (rule tiers instead of exact match)
- `jpackage` installer (bundles its own JRE)
