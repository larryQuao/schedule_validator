package com.schedulevalidator.convert;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.model.ValidationIssue;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;

/**
 * Produces the two workbook outputs:
 *  1. writeConvertedSchedule - an uploaded PDF/image/text report re-created as .xlsx in the
 *     exact layout of the reference sample (header block, labeled columns, totals row).
 *  2. writeValidationReport - findings workbook with a summary sheet and a filterable list.
 */
public final class XlsxExporter {

    // column widths copied from the reference sample
    private static final float[] WIDTHS = {4.14f, 14f, 16.71f, 16.86f, 16f, 14f, 14f, 10.57f, 12f};

    public void writeConvertedSchedule(ContributionReport report, Path out) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            String sheetName = report.getMonthlySheets().isEmpty() ? "CONVERTED"
                    : sanitize(report.getMonthlySheets().get(0).periodLabel());
            Sheet sheet = wb.createSheet(sheetName.length() > 31 ? sheetName.substring(0, 31) : sheetName);

            CellStyle bold = boldStyle(wb);
            CellStyle boldBoxed = boxed(bold);
            CellStyle number = numberStyle(wb);
            CellStyle numberBoxed = boxed(number);

            Row r1 = sheet.createRow(0);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2)); // B1:C1 as in the sample

            Row r2 = sheet.createRow(1);
            text(r2, 1, "SSNIT REGISTRATION NUMBER", boldBoxed);
            text(r2, 3, "NAME OF EMPLOYER:", bold);
            text(r2, 4, nz(report.getEmployerName()), bold);
            text(r2, 7, "NAME OF SCHEME:", bold);
            text(r2, 8, nz(report.getSchemeName()), bold);

            Row r3 = sheet.createRow(2);
            String[] headers = {"#", "Member Code", "SS No", "Surname", "Firstname",
                    "Other Names", "", "Basic Salary", "5% Contribution"};
            for (int i = 0; i < headers.length; i++) {
                text(r3, i, headers[i], boldBoxed);
            }

            List<ScheduleSheet> sheets = report.getMonthlySheets();
            int rowIdx = 3;
            int seq = 1;
            double salaryTotal = 0;
            double contributionTotal = 0;
            for (ScheduleSheet s : sheets) {
                for (ContributionEntry e : s.entries()) {
                    Row row = sheet.createRow(rowIdx++);
                    num(row, 0, seq++, numberBoxed);
                    text(row, 1, e.memberCode(), plain(wb));
                    text(row, 2, e.ssNumber(), plain(wb));
                    text(row, 3, e.surname(), plain(wb));
                    text(row, 4, e.firstName(), plain(wb));
                    text(row, 5, e.otherNames(), plain(wb));
                    text(row, 6, fullNameOf(e), plain(wb));
                    if (e.basicSalary() != null) {
                        num(row, 7, e.basicSalary(), numberBoxed);
                        salaryTotal += e.basicSalary();
                    }
                    if (e.contribution() != null) {
                        num(row, 8, e.contribution(), numberBoxed);
                        contributionTotal += e.contribution();
                    }
                }
            }

            Row totals = sheet.createRow(rowIdx);
            num(totals, 7, salaryTotal, numberBoxed);
            num(totals, 8, contributionTotal, numberBoxed);

            for (int i = 0; i < WIDTHS.length; i++) {
                sheet.setColumnWidth(i, (int) (WIDTHS[i] * 256));
            }
            sheet.createFreezePane(0, 3);

            try (OutputStream os = Files.newOutputStream(out)) {
                wb.write(os);
            }
        }
    }

    /**
     * The single deliverable for the new-month flow: one workbook with the validation
     * summary, the final merged records (NEW = green, REMOVED = red) and all findings.
     */
    public void writeFinalValidationReport(List<com.schedulevalidator.convert.SheetMerger.RecordRow> records,
                                           List<ValidationIssue> issues,
                                           Path refPath, String templateSheet, Path uploadPath,
                                           String newSheetName, int carried, int added, int removed,
                                           Path out) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellStyle bold = boldStyle(wb);
            int r;

            // ---- Summary
            Sheet summary = wb.createSheet("Summary");
            r = infoLine(summary, 0, "Schedule Validator - final validation", "", bold);
            r = infoLine(summary, r, "Generated", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm")));
            r = infoLine(summary, r, "Reference file", refPath == null ? "-" : refPath.toString());
            r = infoLine(summary, r, "Reference sheet (template)", nz(templateSheet));
            r = infoLine(summary, r, "Uploaded file", uploadPath == null ? "-" : uploadPath.toString());
            r = infoLine(summary, r, "New sheet created", nz(newSheetName));
            r++;
            r = infoLine(summary, r, "Records carried over", String.valueOf(carried), bold);
            r = infoLine(summary, r, "New records (green)", String.valueOf(added), bold);
            r = infoLine(summary, r, "Removed records (red)", String.valueOf(removed), bold);
            r++;
            Map<ValidationIssue.Severity, Long> bySeverity = new EnumMap<>(ValidationIssue.Severity.class);
            issues.forEach(i -> bySeverity.merge(i.severity(), 1L, Long::sum));
            r = infoLine(summary, r, "Errors", String.valueOf(bySeverity.getOrDefault(ValidationIssue.Severity.ERROR, 0L)), bold);
            r = infoLine(summary, r, "Warnings", String.valueOf(bySeverity.getOrDefault(ValidationIssue.Severity.WARNING, 0L)), bold);
            r = infoLine(summary, r, "Info", String.valueOf(bySeverity.getOrDefault(ValidationIssue.Severity.INFO, 0L)), bold);
            summary.setColumnWidth(0, 30 * 256);
            summary.setColumnWidth(1, 110 * 256);

            // ---- Updated Records
            Sheet rec = wb.createSheet("Updated Records");
            CellStyle head = boxed(bold);
            String[] headers = {"Status", "S/N", "Member Code", "SS No", "Surname",
                    "Firstname", "Other Names", "Basic Salary", "5% Contribution"};
            Row h = rec.createRow(0);
            for (int i = 0; i < headers.length; i++) text(h, i, headers[i], head);
            CellStyle green = filled(wb, IndexedColors.LIGHT_GREEN.getIndex());
            CellStyle red = filled(wb, IndexedColors.ROSE.getIndex());
            CellStyle num = numberStyle(wb);
            double salaryTotal = 0;
            double contributionTotal = 0;
            int row = 1;
            for (var rec0 : records) {
                Row rr = rec.createRow(row++);
                CellStyle statusStyle = SheetMerger.STATUS_NEW.equals(rec0.status()) ? green
                        : SheetMerger.STATUS_REMOVED.equals(rec0.status()) ? red : null;
                text(rr, 0, rec0.status(), statusStyle);
                if (rec0.seq() != null) num(rr, 1, rec0.seq(), null);
                text(rr, 2, rec0.memberCode(), statusStyle);
                text(rr, 3, rec0.ssNumber(), statusStyle);
                text(rr, 4, rec0.surname(), statusStyle);
                text(rr, 5, rec0.firstName(), statusStyle);
                text(rr, 6, rec0.otherNames(), statusStyle);
                if (rec0.basicSalary() != null) {
                    num(rr, 7, rec0.basicSalary(), statusStyle == null ? num : statusStyle);
                }
                if (rec0.contribution() != null) {
                    num(rr, 8, rec0.contribution(), statusStyle == null ? num : statusStyle);
                }
                if (!SheetMerger.STATUS_REMOVED.equals(rec0.status())) {
                    if (rec0.basicSalary() != null) salaryTotal += rec0.basicSalary();
                    if (rec0.contribution() != null) contributionTotal += rec0.contribution();
                }
            }
            Row tot = rec.createRow(row);
            text(tot, 0, "TOTAL (excl. removed)", bold);
            num(tot, 7, salaryTotal, num);
            num(tot, 8, contributionTotal, num);
            rec.setAutoFilter(new CellRangeAddress(0, Math.max(row - 1, 0), 0, headers.length - 1));
            rec.createFreezePane(0, 1);
            float[] w = {16, 6, 20, 20, 24, 20, 22, 14, 16};
            for (int i = 0; i < w.length; i++) rec.setColumnWidth(i, (int) (w[i] * 256));

            // ---- Findings
            writeFindingsSheet(wb, issues);

            try (OutputStream os = Files.newOutputStream(out)) {
                wb.write(os);
            }
        }
    }

    /** Writes the filterable findings sheet into an open workbook. */
    private void writeFindingsSheet(XSSFWorkbook wb, List<ValidationIssue> issues) {
        Sheet findings = wb.createSheet("Findings");
        String[] headers = {"Severity", "Source", "Sheet", "Location", "Rule", "Message"};
        Row h = findings.createRow(0);
        CellStyle head = boxed(boldStyle(wb));
        for (int i = 0; i < headers.length; i++) text(h, i, headers[i], head);

        CellStyle errorStyle = filled(wb, IndexedColors.ROSE.getIndex());
        CellStyle warnStyle = filled(wb, IndexedColors.LEMON_CHIFFON.getIndex());
        CellStyle infoStyle = filled(wb, IndexedColors.GREY_25_PERCENT.getIndex());
        int row = 1;
        for (ValidationIssue issue : issues) {
            Row rr = findings.createRow(row++);
            text(rr, 0, issue.severity().toString(),
                    styleFor(issue.severity(), errorStyle, warnStyle, infoStyle));
            text(rr, 1, issue.source(), null);
            text(rr, 2, issue.sheet(), null);
            text(rr, 3, issue.location(), null);
            text(rr, 4, issue.rule(), null);
            text(rr, 5, issue.message(), null);
        }
        findings.setAutoFilter(new CellRangeAddress(0, Math.max(row - 1, 0), 0, headers.length - 1));
        findings.createFreezePane(0, 1);
        float[] w = {10, 14, 22, 16, 26, 110};
        for (int i = 0; i < w.length; i++) findings.setColumnWidth(i, (int) (w[i] * 256));
    }

    public void writeValidationReport(List<ValidationIssue> issues, Path refPath, Path uploadPath,
                                      Path out) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet summary = wb.createSheet("Summary");
            CellStyle bold = boldStyle(wb);
            int r = 0;
            r = infoLine(summary, r, "Results", "Schedule Validator", bold);
            r = infoLine(summary, r, "Generated", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm")));
            r = infoLine(summary, r, "Reference file", refPath == null ? "-" : refPath.toString());
            r = infoLine(summary, r, "Uploaded file", uploadPath == null ? "-" : uploadPath.toString());
            r++;
            Map<ValidationIssue.Severity, Long> bySeverity = new EnumMap<>(ValidationIssue.Severity.class);
            issues.forEach(i -> bySeverity.merge(i.severity(), 1L, Long::sum));
            r = infoLine(summary, r, "Errors", String.valueOf(bySeverity.getOrDefault(ValidationIssue.Severity.ERROR, 0L)), bold);
            r = infoLine(summary, r, "Warnings", String.valueOf(bySeverity.getOrDefault(ValidationIssue.Severity.WARNING, 0L)), bold);
            r = infoLine(summary, r, "Info", String.valueOf(bySeverity.getOrDefault(ValidationIssue.Severity.INFO, 0L)), bold);

            Sheet findings = wb.createSheet("Findings");
            String[] headers = {"Severity", "Source", "Sheet", "Location", "Rule", "Message"};
            Row h = findings.createRow(0);
            CellStyle head = boxed(bold);
            for (int i = 0; i < headers.length; i++) text(h, i, headers[i], head);

            CellStyle errorStyle = filled(wb, IndexedColors.ROSE.getIndex());
            CellStyle warnStyle = filled(wb, IndexedColors.LEMON_CHIFFON.getIndex());
            CellStyle infoStyle = filled(wb, IndexedColors.GREY_25_PERCENT.getIndex());
            int row = 1;
            for (ValidationIssue issue : issues) {
                Row rr = findings.createRow(row++);
                text(rr, 0, issue.severity().toString(), styleFor(issue.severity(), errorStyle, warnStyle, infoStyle));
                text(rr, 1, issue.source(), null);
                text(rr, 2, issue.sheet(), null);
                text(rr, 3, issue.location(), null);
                text(rr, 4, issue.rule(), null);
                text(rr, 5, issue.message(), null);
            }
            findings.setAutoFilter(new CellRangeAddress(0, Math.max(row - 1, 0), 0, headers.length - 1));
            findings.createFreezePane(0, 1);
            float[] w = {10, 14, 22, 16, 26, 110};
            for (int i = 0; i < w.length; i++) findings.setColumnWidth(i, (int) (w[i] * 256));

            try (OutputStream os = Files.newOutputStream(out)) {
                wb.write(os);
            }
        }
    }

    // ------------------------------------------------------------------ helpers

    private String fullNameOf(ContributionEntry e) {
        if (e.fullName() != null && !e.fullName().isBlank()) return e.fullName();
        StringBuilder sb = new StringBuilder();
        for (String p : new String[]{e.surname(), e.firstName(), e.otherNames()}) {
            if (p != null && !p.isBlank()) {
                if (!sb.isEmpty()) sb.append(' ');
                sb.append(p.trim());
            }
        }
        return sb.toString();
    }

    private static String sanitize(String name) {
        return name == null ? "CONVERTED" : name.replaceAll("[\\\\/*?:\\[\\]]", "_").trim();
    }

    private static String nz(String s) { return s == null ? "" : s; }

    private int infoLine(Sheet sheet, int row, String label, String value) {
        return infoLine(sheet, row, label, value, null);
    }

    private int infoLine(Sheet sheet, int row, String label, String value, CellStyle style) {
        Row r = sheet.createRow(row);
        text(r, 0, label, style);
        text(r, 1, value, style);
        return row + 1;
    }

    private CellStyle styleFor(ValidationIssue.Severity sev, CellStyle error, CellStyle warn, CellStyle info) {
        return switch (sev) {
            case ERROR -> error;
            case WARNING -> warn;
            case INFO -> info;
        };
    }

    private CellStyle boldStyle(org.apache.poi.ss.usermodel.Workbook wb) {
        Font f = wb.createFont();
        f.setBold(true);
        CellStyle s = wb.createCellStyle();
        s.setFont(f);
        return s;
    }

    private CellStyle plain(org.apache.poi.ss.usermodel.Workbook wb) {
        return wb.createCellStyle();
    }

    private CellStyle boxed(CellStyle base) {
        base.setBorderTop(BorderStyle.THIN);
        base.setBorderBottom(BorderStyle.THIN);
        base.setBorderLeft(BorderStyle.THIN);
        base.setBorderRight(BorderStyle.THIN);
        return base;
    }

    private CellStyle numberStyle(org.apache.poi.ss.usermodel.Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        return s;
    }

    private CellStyle filled(org.apache.poi.ss.usermodel.Workbook wb, short color) {
        CellStyle s = wb.createCellStyle();
        s.setFillForegroundColor(color);
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return s;
    }

    private void text(Row row, int col, String value, CellStyle style) {
        Cell c = row.createCell(col);
        c.setCellValue(value == null ? "" : value);
        if (style != null) c.setCellStyle(style);
    }

    private void num(Row row, int col, double value, CellStyle style) {
        Cell c = row.createCell(col);
        c.setCellValue(value);
        if (style != null) c.setCellStyle(style);
    }
}
