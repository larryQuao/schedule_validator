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
