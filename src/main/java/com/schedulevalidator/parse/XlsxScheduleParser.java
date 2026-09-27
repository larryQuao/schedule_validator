package com.schedulevalidator.parse;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Parses the employer contribution-report workbook layout observed in the reference sample:
 *
 *   row 1..2   header block (SSNIT REGISTRATION NUMBER, NAME OF EMPLOYER:, NAME OF SCHEME:)
 *   row 3      column headers (Member Code | SS No | Surname | Firstname | Other Names | ... |
 *              Basic Salary | 5% Contribution)
 *   row 4..n   member rows, optionally preceded by an unlabeled sequence column
 *   totals row numeric-only row (salary + contribution, no identity cells)
 *
 * Columns are mapped by HEADER TEXT, not position, because the layout shifts between years
 * (e.g. AUGUST 2017 keeps member codes in column B, AUGUST 2026 in column B with the full name
 * in an unlabeled column G). Older sheets carry a second side-by-side table further right; it is
 * excluded by cutting the column map after a gap of unmapped columns.
 */
public final class XlsxScheduleParser {

    private static final double AMOUNT_EPSILON = 0.011; // one cent of float slack
    private static final int MAX_COLUMN_GAP = 2;        // unmapped columns tolerated inside a table
    private static final int MAX_BLANK_RUN = 12;        // blank rows before scanning gives up

    public enum Col { SEQ, MEMBER_CODE, SS_NO, SURNAME, FIRST_NAME, OTHER_NAMES, FULL_NAME,
                      BASIC_SALARY, CONTRIBUTION, TOTAL, DESCRIPTION, DOB }

    /** Public view of a sheet's detected layout, for template-driven sheet building. */
    public record TemplateInfo(int headerRow0, int firstDataRow0, int seqCol0,
                               Map<Integer, Col> columns, int lastDataCol0) {}

    /**
     * Inspects a sheet's layout without parsing its rows: header row index (0-based), first
     * data row, sequence column, mapped roles per column, and the last column that belongs
     * to the main table (side tables are excluded).
     */
    public TemplateInfo inspectTemplate(Sheet sheet) {
        ColumnMap cm = findColumnMap(sheet);
        if (cm == null) return null;
        int lastData = -1;
        for (Map.Entry<Integer, Col> e : cm.byColumn.entrySet()) {
            if (e.getValue() != Col.SEQ && e.getKey() > lastData) lastData = e.getKey();
        }
        return new TemplateInfo(cm.headerRow - 1, cm.headerRow, findSeqColumn(sheet, cm),
                new TreeMap<>(cm.byColumn), lastData);
    }

    /** Parses exactly one sheet into a {@link ScheduleSheet}, or null if no schedule header. */
    public ScheduleSheet parseSingleSheet(Sheet sheet) {
        ColumnMap cm = findColumnMap(sheet);
        if (cm == null) return null;
        ParsedRows rows = scanRows(sheet, cm);
        Period period = parsePeriod(sheet.getSheetName());
        return new ScheduleSheet(sheet.getSheetName(), period.start(), period.end(), period.label(),
                cm.headerRow, rows.entries, rows.statedSalary, rows.statedContribution,
                rows.totalRow, rows.postTotalRows, findMonthLabel(sheet));
    }

    private static final Map<String, Col> HEADER_ROLES = new HashMap<>();
    static {
        HEADER_ROLES.put("MEMBERCODE", Col.MEMBER_CODE);
        HEADER_ROLES.put("SSNO", Col.SS_NO);
        HEADER_ROLES.put("SSNITNO", Col.SS_NO);
        HEADER_ROLES.put("SSNUMBER", Col.SS_NO);
        HEADER_ROLES.put("SURNAME", Col.SURNAME);
        HEADER_ROLES.put("FIRSTNAME", Col.FIRST_NAME);
        HEADER_ROLES.put("OTHERNAMES", Col.OTHER_NAMES);
        HEADER_ROLES.put("BASICSALARY", Col.BASIC_SALARY);
        HEADER_ROLES.put("SALARY", Col.BASIC_SALARY);
        HEADER_ROLES.put("CONTRIBUTION", Col.CONTRIBUTION);
        HEADER_ROLES.put("5%CONTRIBUTION", Col.CONTRIBUTION);
        HEADER_ROLES.put("5%OFBASIC", Col.CONTRIBUTION);
        HEADER_ROLES.put("TOTAL", Col.TOTAL);
        HEADER_ROLES.put("DESCRIPTION", Col.DESCRIPTION);
        HEADER_ROLES.put("DATEOFBIRTH", Col.DOB);
        HEADER_ROLES.put("SN", Col.SEQ);
    }

    private final DataFormatter formatter = new DataFormatter();

    public ContributionReport parse(Path xlsxFile) throws IOException {
        ContributionReport report = new ContributionReport(xlsxFile, ContributionReport.SourceKind.XLSX);
        try (InputStream in = Files.newInputStream(xlsxFile); Workbook wb = WorkbookFactory.create(in)) {
            wb.sheetIterator().forEachRemaining(sheet -> parseSheet(report, sheet));
        }
        report.getMonthlySheets().sort((a, b) -> {
            if (a.periodStart() == null && b.periodStart() == null) return 0;
            if (a.periodStart() == null) return 1;
            if (b.periodStart() == null) return -1;
            return a.periodStart().compareTo(b.periodStart());
        });
        return report;
    }

    // ---------------------------------------------------------------- sheet dispatch

    private void parseSheet(ContributionReport report, Sheet sheet) {
        String name = sheet.getSheetName().trim().toUpperCase(Locale.ROOT);
        if (name.contains("SUSPENSE")) {
            parseSuspense(report, sheet);
            return;
        }
        ColumnMap cm = findColumnMap(sheet);
        if (cm == null) {
            report.addParseNote("Sheet '" + sheet.getSheetName() + "': no schedule header row found, skipped.");
            return;
        }
        ParsedRows rows = scanRows(sheet, cm);
        Period period = parsePeriod(sheet.getSheetName());

        if (name.equals("DETAILS")) {
            // DETAILS is a member x month matrix: contribution is the TOTAL column.
            for (ContributionEntry e : rows.entries) {
                report.getSummaryEntries().add(e);
            }
        } else if (name.equals("SUMMARY")) {
            report.getSummaryEntries().addAll(rows.entries);
        } else {
            report.getMonthlySheets().add(new ScheduleSheet(
                    sheet.getSheetName(),
                    period.start, period.end, period.label,
                    cm.headerRow, rows.entries, rows.statedSalary, rows.statedContribution,
                    rows.totalRow, rows.postTotalRows, findMonthLabel(sheet)));
        }
        captureReportHeader(report, sheet);
    }

    private void parseSuspense(ContributionReport report, Sheet sheet) {
        boolean seenAny = false;
        for (Row row : sheet) {
            List<String> texts = new ArrayList<>();
            Double amount = null;
            String description = null;
            boolean amountSeen = false;
            for (Cell cell : row) {
                String t = text(cell);
                Double n = numeric(cell);
                if (n != null && !isSeqLike(n, cell)) {
                    if (!amountSeen) { amount = n; amountSeen = true; }
                } else if (!t.isBlank()) {
                    if (amountSeen) {
                        if (description == null) description = t.trim();
                    } else {
                        texts.add(t.trim());
                    }
                }
            }
            String payer = String.join(" ", texts).trim();
            if (!payer.isBlank() && amount != null) {
                report.getSuspenseRecords().add(new ContributionReport.SuspenseRecord(
                        row.getRowNum() + 1, payer, amount, description));
                seenAny = true;
            }
        }
        if (!seenAny) {
            report.addParseNote("Suspense sheet '" + sheet.getSheetName() + "' had no recognisable records.");
        }
    }

    private boolean isSeqLike(double n, Cell cell) {
        // sequence numbers / date serials live in narrow columns; amounts live in wide ones.
        return n == Math.floor(n) && (cell.getColumnIndex() <= 1
                || (n > 40000 && n < 60000)); // excel date serials
    }

    // ---------------------------------------------------------------- header mapping

    private static final class ColumnMap {
        int headerRow;
        final Map<Integer, Col> byColumn = new TreeMap<>();
        int firstMapped = -1;
        int lastMapped = -1;

        Col roleOf(int col) { return byColumn.get(col); }
    }

    private ColumnMap findColumnMap(Sheet sheet) {
        for (int r = 0; r < Math.min(15, sheet.getLastRowNum() + 1); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            ColumnMap cm = mapHeaderRow(row);
            if (cm != null) {
                cm.headerRow = r + 1; // 1-based row number of the column headers
                return cm;
            }
        }
        return null;
    }

    private ColumnMap mapHeaderRow(Row row) {
        ColumnMap cm = new ColumnMap();
        int unmappedRun = 0;
        boolean tableOpen = false;
        for (Cell cell : row) {
            String norm = normalize(text(cell));
            Col role = norm.isEmpty() ? null : HEADER_ROLES.get(norm);
            if (role == null) {
                // "5% Contribution" normalizes to "5%CONTRIBUTION" already; try contains fallback
                if (norm.contains("CONTRIBUTION") && !norm.contains("TOTAL")) role = Col.CONTRIBUTION;
                else if (norm.contains("SURNAME")) role = Col.SURNAME;
                else if (norm.contains("SS") && norm.contains("NO")) role = Col.SS_NO;
                else if (norm.contains("MEMBER")) role = Col.MEMBER_CODE;
                else if (norm.contains("BASIC")) role = Col.BASIC_SALARY;
            }
            if (role == null) {
                unmappedRun++;
                if (tableOpen && unmappedRun > MAX_COLUMN_GAP) break; // side table starts
            } else {
                if (cm.firstMapped == -1) cm.firstMapped = cell.getColumnIndex();
                if (!cm.byColumn.containsValue(role)) {
                    cm.byColumn.put(cell.getColumnIndex(), role);
                    cm.lastMapped = cell.getColumnIndex();
                }
                tableOpen = true;
                unmappedRun = 0;
            }
        }
        if (cm.firstMapped == -1 || !cm.byColumn.containsValue(Col.SS_NO)
                || (!cm.byColumn.containsValue(Col.BASIC_SALARY) && !cm.byColumn.containsValue(Col.CONTRIBUTION))) {
            return null; // not a schedule header
        }
        return cm;
    }

    // ---------------------------------------------------------------- row scanning

    private static final class ParsedRows {
        final List<ContributionEntry> entries = new ArrayList<>();
        Double statedSalary;
        Double statedContribution;
        int totalRow;
        final List<Integer> postTotalRows = new ArrayList<>();
    }

    private ParsedRows scanRows(Sheet sheet, ColumnMap cm) {
        ParsedRows out = new ParsedRows();
        int seqCol = findSeqColumn(sheet, cm);
        int blankRun = 0;
        boolean seenEntry = false;

        for (int r = cm.headerRow; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) { blankRun++; if (seenEntry && blankRun > MAX_BLANK_RUN) break; continue; }

            Integer seq = seqCol >= 0 ? intOrNull(cellAt(row, seqCol)) : null;
            String memberCode = textAt(row, cm, Col.MEMBER_CODE);
            String ss = textAt(row, cm, Col.SS_NO);
            String surname = textAt(row, cm, Col.SURNAME);
            String first = textAt(row, cm, Col.FIRST_NAME);
            String other = textAt(row, cm, Col.OTHER_NAMES);
            String full = textAt(row, cm, Col.FULL_NAME);
            Double salary = dblAt(row, cm, Col.BASIC_SALARY);
            Double contrib = dblAt(row, cm, Col.CONTRIBUTION);
            Double total = dblAt(row, cm, Col.TOTAL);
            String descr = textAt(row, cm, Col.DESCRIPTION);
            String dob = textAt(row, cm, Col.DOB);

            boolean hasIdentity = notBlank(ss) || notBlank(surname) || notBlank(first) || notBlank(full) || notBlank(memberCode);
            boolean hasAmount = salary != null || contrib != null || total != null;

            if (hasIdentity && hasAmount) {
                int rowNo = r + 1;
                if (out.totalRow > 0) {
                    out.postTotalRows.add(rowNo); // entries appended below the totals row
                }
                out.entries.add(new ContributionEntry(rowNo, seq, blank(memberCode), blank(ss),
                        blank(surname), blank(first), blank(other), blank(full), salary,
                        total != null && contrib == null ? total : contrib, blank(descr), blank(dob)));
                seenEntry = true;
                blankRun = 0;
            } else if (!hasIdentity && seenEntry && salary != null && contrib != null
                    && Math.abs(salary) > AMOUNT_EPSILON && Math.abs(contrib) > AMOUNT_EPSILON
                    && (out.statedSalary == null && out.statedContribution == null)) {
                // numeric-only row after entries: the printed totals row (only the first counts)
                out.statedSalary = salary;
                out.statedContribution = contrib;
                out.totalRow = r + 1;
                blankRun = 0;
            } else {
                blankRun++;
                if (seenEntry && blankRun > MAX_BLANK_RUN) break;
            }
        }
        return out;
    }

    private int findSeqColumn(Sheet sheet, ColumnMap cm) {
        Row firstData = sheet.getRow(cm.headerRow); // header row itself holds labels; check it and next
        Row candidate = sheet.getRow(cm.headerRow + 1);
        for (int c = 0; candidate != null && c < Math.max(cm.firstMapped, 0); c++) {
            if (cm.byColumn.containsKey(c)) continue;
            Integer v = intOrNull(cellAt(candidate, c));
            if (v != null && v == 1) {
                // confirm it keeps counting on the next row when possible
                Row next = sheet.getRow(cm.headerRow + 2);
                if (next == null) return c;
                Integer v2 = intOrNull(cellAt(next, c));
                if (v2 == null || v2 == 2 || v2 > v) return c;
            }
        }
        return -1;
    }

    // ---------------------------------------------------------------- header block + labels

    private void captureReportHeader(ContributionReport report, Sheet sheet) {
        if (report.getEmployerName() != null && report.getSchemeName() != null) return;
        for (int r = 0; r < Math.min(5, sheet.getLastRowNum() + 1); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (Cell cell : row) {
                String t = text(cell).toUpperCase(Locale.ROOT);
                if (t.contains("NAME OF EMPLOYER")) {
                    report.setEmployerName(valueRightOf(row, cell));
                } else if (t.contains("NAME OF SCHEME") || t.startsWith("SCHEME:")) {
                    report.setSchemeName(valueRightOf(row, cell));
                } else if (t.contains("SSNIT REGISTRATION")) {
                    report.setSsnitRegNumber(valueRightOf(row, cell));
                }
            }
        }
    }

    private String valueRightOf(Row row, Cell label) {
        for (int c = label.getColumnIndex() + 1; c <= label.getColumnIndex() + 3; c++) {
            String t = text(cellAt(row, c));
            if (notBlank(t)) return t.trim();
        }
        return null;
    }

    private String findMonthLabel(Sheet sheet) {
        for (int r = 0; r < Math.min(5, sheet.getLastRowNum() + 1); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (Cell cell : row) {
                String t = text(cell).toUpperCase(Locale.ROOT);
                Matcher m = Pattern.compile("MONTH\\s*:?\\s*([A-Z]{3,9})").matcher(t);
                if (m.find()) return m.group(1);
            }
        }
        return null;
    }

    // ---------------------------------------------------------------- period parsing

    private record Period(String start, String end, String label) {}

    private static final Pattern MONTH_WORD = Pattern.compile(
            "(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*");
    private static final Pattern YEAR4 = Pattern.compile("\\b(\\d{4})\\b");
    private static final Pattern YEAR2 = Pattern.compile("(?:^|[-\\s])(\\d{2})\\b");

    /** Best-effort period extraction from sheet names like "AUGUST 2017", "MAR-20 TO JUL-20",
     *  "OCTOBER TO DECEMBER 2019", "DEC-20 & FEB-21", "jan-jun 13 and penalty". */
    static Period parsePeriod(String sheetName) {
        String name = sheetName.toUpperCase(Locale.ROOT);
        Matcher mm = MONTH_WORD.matcher(name);
        List<int[]> months = new ArrayList<>(); // {monthNum, position}
        while (mm.find()) months.add(new int[]{monthNumber(mm.group(1)), mm.start()});
        if (months.isEmpty()) return new Period(null, null, sheetName);

        List<Integer> years = new ArrayList<>();
        Matcher y4 = YEAR4.matcher(name);
        while (y4.find()) years.add(Integer.parseInt(y4.group(1)));
        if (years.isEmpty()) {
            Matcher y2 = YEAR2.matcher(name);
            while (y2.find()) years.add(2000 + Integer.parseInt(y2.group(1)));
        }

        int n = months.size();
        String start = null;
        String end = null;
        for (int i = 0; i < n; i++) {
            int year;
            if (years.size() == n) year = years.get(i);
            else if (!years.isEmpty()) year = (i == 0) ? years.get(0) : years.get(Math.min(i, years.size() - 1));
            else return new Period(null, null, sheetName);
            String ym = String.format("%04d-%02d", year, months.get(i)[0]);
            if (start == null) start = ym;
            end = ym;
        }
        if (start.equals(end)) end = null;
        return new Period(start, end, sheetName);
    }

    private static int monthNumber(String abbrev) {
        return switch (abbrev) {
            case "JAN" -> 1; case "FEB" -> 2; case "MAR" -> 3; case "APR" -> 4;
            case "MAY" -> 5; case "JUN" -> 6; case "JUL" -> 7; case "AUG" -> 8;
            case "SEP" -> 9; case "OCT" -> 10; case "NOV" -> 11; case "DEC" -> 12;
            default -> 0;
        };
    }

    // ---------------------------------------------------------------- cell helpers

    /** Uppercase, strip everything but A-Z0-9 so header text matches regardless of spacing/case. */
    private static String normalize(String s) {
        return s == null ? "" : s.toUpperCase(Locale.ROOT).replaceAll("[^A-Z0-9%]", "");
    }

    private Cell cellAt(Row row, int col) {
        return col < 0 ? null : row.getCell(col);
    }

    private String text(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.NUMERIC) {
            double v = cell.getNumericCellValue();
            return v == Math.floor(v) ? String.valueOf((long) v) : String.valueOf(v);
        }
        return formatter.formatCellValue(cell).trim();
    }

    private String textAt(Row row, ColumnMap cm, Col role) {
        for (Map.Entry<Integer, Col> e : cm.byColumn.entrySet()) {
            if (e.getValue() == role) return text(cellAt(row, e.getKey()));
        }
        return "";
    }

    private Double dblAt(Row row, ColumnMap cm, Col role) {
        for (Map.Entry<Integer, Col> e : cm.byColumn.entrySet()) {
            if (e.getValue() == role) return numeric(cellAt(row, e.getKey()));
        }
        return null;
    }

    private Double numeric(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
        if (cell.getCellType() == CellType.FORMULA) {
            try {
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) return cell.getNumericCellValue();
            } catch (Exception ignored) { }
            String s = text(cell);
            try { return s.isBlank() ? null : Double.parseDouble(s.replace(",", "")); }
            catch (NumberFormatException e) { return null; }
        }
        if (cell.getCellType() == CellType.STRING) {
            String s = cell.getStringCellValue().replace(",", "").trim();
            if (s.isEmpty()) return null;
            try { return Double.parseDouble(s); } catch (NumberFormatException e) { return null; }
        }
        return null;
    }

    private Integer intOrNull(Cell cell) {
        Double d = numeric(cell);
        return (d != null && d == Math.floor(d)) ? (int) (double) d : null;
    }

    private static String blank(String s) { return s == null ? "" : s; }
    private static boolean notBlank(String s) { return s != null && !s.isBlank(); }
}
