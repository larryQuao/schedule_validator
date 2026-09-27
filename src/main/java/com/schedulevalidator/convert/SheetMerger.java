package com.schedulevalidator.convert;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.parse.XlsxScheduleParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Builds the "new month" sheet inside a copy of the contribution report.
 *
 * The selected existing sheet acts as the template: header block, column styles, widths,
 * merged regions and per-cell formulas (row-shifted, e.g. {@code =0.05*H4}) are carried
 * into the new sheet. Records found in both the template and the upload are written with
 * the upload's (new-month) values; records only in the upload are appended and highlighted
 * GREEN; records of the template missing from the upload are marked REMOVED, written below
 * the totals row and highlighted RED — the totals (=SUM over kept + new rows only) exclude
 * them. The original workbook file is never modified; the result is a new file.
 */
public final class SheetMerger {

    public record MergeSummary(Path updatedWorkbook, String newSheetName,
                               int matched, int added, int removed, List<String> notes) {}

    private static final Pattern A1_REF = Pattern.compile("(\\$?[A-Z]{1,3})(\\$?)(\\d{1,5})");

    public MergeSummary build(Path referenceXlsx, String templateSheetName,
                              ScheduleSheet uploadSheet, List<ContributionEntry> uploadEntries,
                              Path outDir) throws IOException {
        final XSSFWorkbook wb;
        try {
            wb = new XSSFWorkbook(referenceXlsx.toFile());
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            throw new IOException("Not a valid .xlsx file: " + referenceXlsx, e);
        }
        try (wb) {
            XSSFSheet template = wb.getSheet(templateSheetName);
            if (template == null) {
                throw new IOException("Sheet not found in the reference workbook: " + templateSheetName);
            }
            XlsxScheduleParser parser = new XlsxScheduleParser();
            ScheduleSheet existing = parser.parseSingleSheet(template);
            if (existing == null) {
                throw new IOException("Selected sheet has no recognizable schedule header: " + templateSheetName);
            }
            XlsxScheduleParser.TemplateInfo ti = parser.inspectTemplate(template);
            int dataStart1 = ti.headerRow0() + 2; // 1-based first data row

            // ---------------------------------------------------------------- matching
            Map<String, ContributionEntry> upByKey = new HashMap<>();
            for (ContributionEntry up : uploadEntries) {
                String k = keyOf(up);
                if (k != null) upByKey.putIfAbsent(k, up);
            }
            List<ContributionEntry> matched = new ArrayList<>(); // template order, upload values
            List<ContributionEntry> removed = new ArrayList<>(); // template order
            Set<String> usedKeys = new HashSet<>();
            for (ContributionEntry ex : existing.entries()) {
                String k = keyOf(ex);
                ContributionEntry up = k == null ? null : upByKey.get(k);
                if (up != null) {
                    matched.add(up);
                    usedKeys.add(k);
                } else {
                    removed.add(ex);
                }
            }
            List<ContributionEntry> added = new ArrayList<>(); // upload order
            for (ContributionEntry up : uploadEntries) {
                String k = keyOf(up);
                if (k == null || !usedKeys.contains(k)) added.add(up);
            }

            // ---------------------------------------------------------------- sheet + structure
            String newName = uniqueName(wb, uploadSheet.periodLabel());
            XSSFSheet sh = wb.createSheet(newName);

            for (int r = 0; r <= ti.headerRow0(); r++) { // header block + column headers
                copyRowVerbatim(template, sh, r);
            }
            for (int i = template.getNumMergedRegions() - 1; i >= 0; i--) {
                var range = template.getMergedRegion(i);
                if (range.getLastRow() <= ti.headerRow0()) {
                    sh.addMergedRegion(range);
                }
            }

            // per-column styles come from the template's data rows; formulas are shifted
            CellStyle[] colStyles = new CellStyle[ti.lastDataCol0() + 1];
            String[] colFormulas = new String[ti.lastDataCol0() + 1]; // template formula, if any
            int[] colFormulaRow1 = new int[ti.lastDataCol0() + 1];    // 1-based row it was read from
            for (int c = 0; c <= ti.lastDataCol0(); c++) {
                for (int probe = 0; probe < 6; probe++) {
                    Row trow = template.getRow(dataStart1 - 1 + probe);
                    Cell tc = trow == null ? null : trow.getCell(c);
                    if (tc == null) continue;
                    if (colStyles[c] == null) colStyles[c] = tc.getCellStyle();
                    if (colFormulas[c] == null && tc.getCellType() == CellType.FORMULA) {
                        colFormulas[c] = tc.getCellFormula();
                        colFormulaRow1[c] = tc.getRowIndex() + 1;
                    }
                    if (colStyles[c] != null && colFormulas[c] != null) break;
                }
            }

            Map<Short, CellStyle> greenCache = new HashMap<>();
            Map<Short, CellStyle> redCache = new HashMap<>();
            List<String> notes = new ArrayList<>();

            int seq = 1;
            int r0 = ti.headerRow0() + 1; // 0-based row cursor
            for (ContributionEntry up : matched) {
                writeDataRow(sh, r0++, up, seq++, ti, colStyles, colFormulas, colFormulaRow1, null, null, notes);
            }
            for (ContributionEntry up : added) {
                writeDataRow(sh, r0++, up, seq++, ti, colStyles, colFormulas, colFormulaRow1, greenCache, IndexedColors.LIGHT_GREEN, notes);
            }
            int lastData1 = r0; // 1-based row number of the last data row

            // ---------------------------------------------------------------- totals row
            int rTot0 = r0++;
            Row templateTotals = existing.totalRow() > 0 ? template.getRow(existing.totalRow() - 1) : null;
            Row totals = sh.createRow(rTot0);
            double salarySum = 0;
            double contributionSum = 0;
            for (ContributionEntry e : matched) {
                if (e.basicSalary() != null) salarySum += e.basicSalary();
                if (e.contribution() != null) contributionSum += e.contribution();
            }
            for (ContributionEntry e : added) {
                if (e.basicSalary() != null) salarySum += e.basicSalary();
                if (e.contribution() != null) contributionSum += e.contribution();
            }
            for (int c = 0; c <= ti.lastDataCol0(); c++) {
                Cell tc = templateTotals == null ? null : templateTotals.getCell(c);
                Cell nc = totals.createCell(c);
                if (tc != null) nc.setCellStyle(tc.getCellStyle());
                boolean isAmountCol = c == ti.columns().entrySet().stream()
                        .filter(e -> e.getValue() == XlsxScheduleParser.Col.BASIC_SALARY)
                        .mapToInt(Map.Entry::getKey).findFirst().orElse(-1)
                        || c == ti.columns().entrySet().stream()
                        .filter(e -> e.getValue() == XlsxScheduleParser.Col.CONTRIBUTION)
                        .mapToInt(Map.Entry::getKey).findFirst().orElse(-1);
                if (isAmountCol) {
                    String letter = CellReference.convertNumToColString(c);
                    double fallback = c == colOf(ti, XlsxScheduleParser.Col.BASIC_SALARY) ? salarySum : contributionSum;
                    try {
                        nc.setCellFormula("SUM(" + letter + dataStart1 + ":" + letter + lastData1 + ")");
                    } catch (RuntimeException ex) {
                        nc.setCellValue(fallback);
                        notes.add("Totals formula replaced with computed value " + fallback
                                + " for column " + letter + ": " + ex.getMessage());
                    }
                } else if (tc != null) {
                    copyCellValueSafe(tc, nc, notes);
                }
            }

            // ---------------------------------------------------------------- removed rows
            for (ContributionEntry ex : removed) {
                writeRemovedRow(sh, r0++, ex, ti, colStyles, colFormulas, colFormulaRow1, redCache, notes);
            }

            // ---------------------------------------------------------------- geometry
            for (int c = 0; c <= ti.lastDataCol0(); c++) {
                sh.setColumnWidth(c, template.getColumnWidth(c));
            }
            var pane = template.getPaneInformation();
            if (pane != null && pane.isFreezePane()) {
                sh.createFreezePane(pane.getHorizontalSplitPosition(), pane.getVerticalSplitPosition());
            }

            // ---------------------------------------------------------------- save
            Files.createDirectories(outDir);
            String base = referenceXlsx.getFileName().toString().replaceAll("(?i)\\.xlsx?$", "");
            Path out = outDir.resolve(base + " - updated.xlsx");
            try (var os = Files.newOutputStream(out)) {
                wb.write(os);
            }
            return new MergeSummary(out, newName, matched.size(), added.size(), removed.size(), notes);
        }
    }

    // ------------------------------------------------------------------ row writers

    private void writeDataRow(XSSFSheet sh, int r0, ContributionEntry up, int seq,
                              XlsxScheduleParser.TemplateInfo ti, CellStyle[] colStyles,
                              String[] colFormulas, int[] colFormulaRow1,
                              Map<Short, CellStyle> tintCache, IndexedColors tint, List<String> notes) {
        Row row = sh.createRow(r0);
        if (ti.seqCol0() >= 0) {
            Cell seqCell = row.createCell(ti.seqCol0());
            seqCell.setCellValue(seq);
            if (colStyles[ti.seqCol0()] != null) {
                seqCell.setCellStyle(style(sh.getWorkbook(), colStyles[ti.seqCol0()], tintCache, tint));
            }
        }
        for (var en : ti.columns().entrySet()) {
            int c = en.getKey();
            if (en.getValue() == XlsxScheduleParser.Col.SEQ) continue;
            Cell cell = row.createCell(c);
            if (colStyles[c] != null) cell.setCellStyle(style(sh.getWorkbook(), colStyles[c], tintCache, tint));
            if (colFormulas[c] != null && c != colOf(ti, XlsxScheduleParser.Col.MEMBER_CODE)
                    && c != colOf(ti, XlsxScheduleParser.Col.SS_NO)
                    && tryFormula(cell, colFormulas[c], colFormulaRow1[c], r0 + 1, notes)) {
                continue;
            }
            switch (en.getValue()) {
                case MEMBER_CODE -> setIf(cell, up.memberCode());
                case SS_NO -> setIf(cell, up.ssNumber());
                case SURNAME -> setIf(cell, up.surname());
                case FIRST_NAME -> setIf(cell, up.firstName());
                case OTHER_NAMES -> setIf(cell, up.otherNames());
                case BASIC_SALARY -> { if (up.basicSalary() != null) cell.setCellValue(up.basicSalary()); }
                case CONTRIBUTION -> { if (up.contribution() != null) cell.setCellValue(up.contribution()); }
                case DESCRIPTION -> setIf(cell, up.description());
                case DOB -> setIf(cell, up.dateOfBirth());
                default -> { }
            }
        }
    }

    private void writeRemovedRow(XSSFSheet sh, int r0, ContributionEntry ex,
                                 XlsxScheduleParser.TemplateInfo ti, CellStyle[] colStyles,
                                 String[] colFormulas, int[] colFormulaRow1,
                                 Map<Short, CellStyle> redCache, List<String> notes) {
        Row row = sh.createRow(r0);
        int seqCol = ti.seqCol0();
        int firstCol = seqCol >= 0 ? seqCol : 0;
        for (int c = 0; c <= ti.lastDataCol0(); c++) {
            if (colStyles[c] == null && c != firstCol) continue;
            Cell cell = row.createCell(c);
            if (colStyles[c] != null) cell.setCellStyle(style(sh.getWorkbook(), colStyles[c], redCache, IndexedColors.ROSE));
            if (c == firstCol) {
                cell.setCellValue("REMOVED");
                continue;
            }
            if (colFormulas[c] != null && c != colOf(ti, XlsxScheduleParser.Col.MEMBER_CODE)
                    && c != colOf(ti, XlsxScheduleParser.Col.SS_NO)
                    && tryFormula(cell, colFormulas[c], colFormulaRow1[c], r0 + 1, notes)) {
                continue;
            }
            switch (ti.columns().getOrDefault(c, XlsxScheduleParser.Col.SEQ)) {
                case MEMBER_CODE -> setIf(cell, ex.memberCode());
                case SS_NO -> setIf(cell, ex.ssNumber());
                case SURNAME -> setIf(cell, ex.surname());
                case FIRST_NAME -> setIf(cell, ex.firstName());
                case OTHER_NAMES -> setIf(cell, ex.otherNames());
                case BASIC_SALARY -> { if (ex.basicSalary() != null) cell.setCellValue(ex.basicSalary()); }
                case CONTRIBUTION -> { if (ex.contribution() != null) cell.setCellValue(ex.contribution()); }
                default -> { }
            }
        }
    }

    private void copyRowVerbatim(XSSFSheet from, XSSFSheet to, int r0) {
        Row src = from.getRow(r0);
        if (src == null) return;
        Row dst = to.createRow(r0);
        for (Cell c : src) {
            Cell nc = dst.createCell(c.getColumnIndex());
            nc.setCellStyle(c.getCellStyle());
            copyCellValue(c, nc, 0);
        }
        if (src.getHeight() != 0) dst.setHeight(src.getHeight());
    }

    private static void copyCellValue(Cell src, Cell dst, int rowShift) {
        switch (src.getCellType()) {
            case STRING -> dst.setCellValue(src.getStringCellValue());
            case NUMERIC -> dst.setCellValue(src.getNumericCellValue());
            case BOOLEAN -> dst.setCellValue(src.getBooleanCellValue());
            case FORMULA -> {
                String f = src.getCellFormula();
                dst.setCellFormula(rowShift == 0 ? f : shiftFormula(f, src.getRowIndex() + 1, src.getRowIndex() + 1 + rowShift));
            }
            case ERROR -> dst.setCellErrorValue(src.getErrorCellValue());
            default -> { }
        }
    }

    // ------------------------------------------------------------------ helpers

    private static CellStyle style(XSSFWorkbook wb, CellStyle base, Map<Short, CellStyle> cache, IndexedColors color) {
        if (base == null || color == null) return base;
        return cache.computeIfAbsent(base.getIndex(), k -> {
            CellStyle s = wb.createCellStyle();
            s.cloneStyleFrom(base);
            s.setFillForegroundColor(color.getIndex());
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return s;
        });
    }

    private static int colOf(XlsxScheduleParser.TemplateInfo ti, XlsxScheduleParser.Col role) {
        return ti.columns().entrySet().stream()
                .filter(e -> e.getValue() == role)
                .mapToInt(Map.Entry::getKey).findFirst().orElse(-1);
    }

    private static void setIf(Cell cell, String value) {
        if (value != null && !value.isBlank()) cell.setCellValue(value);
    }

    private static String keyOf(ContributionEntry e) {
        String ss = e.normalizedSs();
        return ss.isEmpty() ? "NAME:" + e.normalizedName() : ss;
    }

    /** Shifts relative A1 row references (H4 stays column H, row follows old→new). */
    static String shiftFormula(String formula, int oldRow1, int newRow1) {
        int delta = newRow1 - oldRow1;
        if (delta == 0) return formula;
        Matcher m = A1_REF.matcher(formula);
        StringBuilder sb = new StringBuilder();
        int last = 0;
        while (m.find()) {
            sb.append(formula, last, m.start());
            if (m.group(2).isEmpty()) { // relative row reference only
                int row = Integer.parseInt(m.group(3));
                sb.append(m.group(1)).append(row + delta);
            } else {
                sb.append(m.group()); // absolute row: keep
            }
            last = m.end();
        }
        sb.append(formula.substring(last));
        return sb.toString();
    }

    /**
     * Attempts to set a row-shifted formula; on any parser rejection writes nothing and
     * records a note so the merge never fails over an exotic template formula.
     */
    private static boolean tryFormula(Cell cell, String original, int oldRow1, int newRow1,
                                      List<String> notes) {
        String shifted = shiftFormula(original, oldRow1, newRow1);
        try {
            cell.setCellFormula(shifted);
            return true;
        } catch (RuntimeException e) {
            notes.add("Row " + (cell.getRowIndex() + 1) + ": formula '" + original
                    + "' could not be carried over (" + e.getMessage() + ") - value written instead.");
            return false;
        }
    }

    private static void copyCellValueSafe(Cell src, Cell dst, List<String> notes) {
        try {
            copyCellValue(src, dst, 0);
        } catch (RuntimeException e) {
            notes.add("Cell " + src.getAddress() + " copied blank: " + e.getMessage());
        }
    }

    private static String uniqueName(XSSFWorkbook wb, String wanted) {
        String base = wanted == null || wanted.isBlank() ? "NEW ENTRIES" : wanted.trim().toUpperCase(Locale.ROOT);
        base = base.replaceAll("[\\\\/*?:\\[\\]]", "_");
        if (base.length() > 28) base = base.substring(0, 28);
        String name = base;
        int i = 2;
        while (wb.getSheet(name) != null) {
            name = base + " " + i++;
        }
        return name;
    }
}
