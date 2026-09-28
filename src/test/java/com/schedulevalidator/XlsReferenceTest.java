package com.schedulevalidator;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ScheduleSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The merge flow must accept a legacy .xls reference workbook: template copy, green new
 * row, red removed row, and an output whose extension matches the input format.
 */
class XlsReferenceTest {

    @TempDir
    Path tempDir;

    @Test
    void mergeAcceptsLegacyXlsReference() throws Exception {
        // build a small .xls workbook in the reference layout
        Path xlsRef = tempDir.resolve("legacy_report.xls");
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            Sheet sh = wb.createSheet("JULY 2026");
            Row h = sh.createRow(2);
            h.createCell(1).setCellValue("Member Code");
            h.createCell(2).setCellValue("SS No");
            h.createCell(3).setCellValue("Surname");
            h.createCell(4).setCellValue("Firstname");
            h.createCell(5).setCellValue("Other Names");
            h.createCell(7).setCellValue("Basic Salary");
            h.createCell(8).setCellValue("5% Contribution");
            String[][] members = {
                    {"GHA-001-1", "C018207011306", "NORTEY", "NORA", "", "1200", "60"},
                    {"GHA-002-2", "B127901300019", "MENSAH", "KENNETH", "", "1000", "50"},
                    {"GHA-003-3", "C099006260026", "LARTEY", "LYDIA", "B.", "700", "35"},
            };
            int r = 3;
            for (String[] m : members) {
                Row row = sh.createRow(r++);
                row.createCell(0).setCellValue(r - 3);
                for (int i = 0; i < m.length; i++) row.createCell(i + 1).setCellValue(m[i]);
            }
            Row tot = sh.createRow(r);
            tot.createCell(7).setCellValue(2900);
            tot.createCell(8).setCellValue(145);
            try (var os = Files.newOutputStream(xlsRef)) {
                wb.write(os);
            }
        }

        // upload: keep 2 members, drop LARTEY, add one new member
        List<ContributionEntry> upload = new ArrayList<>();
        upload.add(new ContributionEntry(1, 1, "GHA-001-1", "C018207011306", "NORTEY", "NORA", "", "", 1300.0, 65.0, "", ""));
        upload.add(new ContributionEntry(2, 2, "GHA-002-2", "B127901300019", "MENSAH", "KENNETH", "", "", 1000.0, 50.0, "", ""));
        upload.add(new ContributionEntry(3, 3, "", "Z000000000001", "ASHONG", "KOFI", "", "", 800.0, 40.0, "", ""));
        ScheduleSheet upSheet = new ScheduleSheet("AUGUST 2026", "2026-08", null, "AUGUST 2026",
                3, upload, null, null, 0, List.of(), null);

        var summary = new com.schedulevalidator.convert.SheetMerger().build(
                xlsRef, "JULY 2026", upSheet, upload, tempDir);

        assertEquals(2, summary.matched());
        assertEquals(1, summary.added());
        assertEquals(1, summary.removed());

        Path out = summary.updatedWorkbook();
        assertTrue(out.getFileName().toString().endsWith(" - updated.xls"),
                "output must match input format: " + out);
        try (Workbook check = org.apache.poi.ss.usermodel.WorkbookFactory.create(out.toFile())) {
            Sheet newsheet = check.getSheet("AUGUST 2026");
            assertTrue(newsheet != null, "new sheet missing in .xls output");
            // 2 carried + 1 new + totals + 1 removed = rows 3..6 data, 7 totals, 8 removed
            assertEquals("REMOVED", newsheet.getRow(7).getCell(0).getStringCellValue());
            assertEquals("ASHONG", newsheet.getRow(5).getCell(3).getStringCellValue());
        }
    }
}
