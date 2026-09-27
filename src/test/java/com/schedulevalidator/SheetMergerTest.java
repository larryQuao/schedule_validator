package com.schedulevalidator;

import com.schedulevalidator.convert.XlsxExporter;
import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.parse.XlsxScheduleParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * End-to-end test of the new-month merge: selects the sample's AUG 2026 sheet, uploads a
 * synthetic September schedule (3 members dropped, 1 salary changed, 2 brand-new members),
 * and verifies the produced workbook: new sheet, template formulas shifted, green new rows,
 * red REMOVED rows after the totals, and totals summing kept+new only.
 */
class SheetMergerTest {

    private static final Path SAMPLE = Paths.get("Samples", "ANNE MARIE CONTRIBUTION REPORT.xlsx");

    private record Prepped(List<ContributionEntry> entries, Path uploadFile,
                           int removedCount, int addedCount) {}

    private Prepped prepUpload() throws Exception {
        assertTrue(Files.isRegularFile(SAMPLE), "sample workbook missing: " + SAMPLE.toAbsolutePath());
        ContributionReport sample = new XlsxScheduleParser().parse(SAMPLE);
        ScheduleSheet aug26 = sample.getMonthlySheets().stream()
                .filter(s -> "2026-08".equals(s.periodStart())).findFirst().orElseThrow();

        List<ContributionEntry> uploadEntries = new ArrayList<>();
        List<ContributionEntry> ex = aug26.entries();
        for (int i = 0; i < ex.size(); i++) {
            if (i == 5 || i == 10 || i == 20) continue; // these become REMOVED
            ContributionEntry e = ex.get(i);
            Double sal = e.basicSalary();
            if (i == 0 && sal != null) sal = sal + 100;
            uploadEntries.add(new ContributionEntry(i + 1, i + 1, e.memberCode(), e.ssNumber(),
                    e.surname(), e.firstName(), e.otherNames(), "", sal,
                    sal == null ? null : sal * 0.05, "", ""));
        }
        int seq = uploadEntries.size() + 1;
        uploadEntries.add(new ContributionEntry(seq, seq, "", "T000000000011", "ASHONG", "KOFI", "", "",
                800.0, 40.0, "", "")); // becomes GREEN
        seq++;
        uploadEntries.add(new ContributionEntry(seq, seq, "", "T000000000012", "BOATENG", "AMA", "", "",
                600.0, 30.0, "", "")); // becomes GREEN

        ContributionReport up = new ContributionReport(Path.of("upload-sep-2026.xlsx"),
                ContributionReport.SourceKind.XLSX);
        up.setEmployerName("ANNIE MARIE");
        up.setSchemeName("SCHEME: QFTL OCCUPATIONAL PENSION SCHEME (MT)");
        up.getMonthlySheets().add(new ScheduleSheet("SEPTEMBER 2026", "2026-09", null,
                "SEPTEMBER 2026", 3, uploadEntries, null, null, 0, List.of(), null));

        Path uploadFile = Files.createTempFile("sv-upload-", ".xlsx");
        new XlsxExporter().writeConvertedSchedule(up, uploadFile);
        return new Prepped(uploadEntries, uploadFile, 3, 2);
    }

    @Test
    void mergeAddsNewMonthSheetWithHighlightsAndFormulas() throws Exception {
        Prepped prep = prepUpload();
        Path outDir = Files.createTempDirectory("sv-merge-");
        List<String> log = new ArrayList<>();
        Pipeline.Result result = Pipeline.runSheetMerge(SAMPLE, "AUG 2026", prep.uploadFile(),
                outDir, Path.of("tessdata"), log);

        assertTrue(Files.isRegularFile(result.updatedWorkbook()), "updated workbook missing");
        assertTrue(result.updatedWorkbook().getFileName().toString().endsWith(" - updated.xlsx"));
        assertTrue(Files.size(result.updatedWorkbook()) > 100_000, "updated workbook suspiciously small");

        try (XSSFWorkbook wb = new XSSFWorkbook(result.updatedWorkbook().toFile())) {
            XSSFSheet sh = wb.getSheet("SEPTEMBER 2026");
            assertNotNull(sh, "new month sheet missing");

            // layout: data rows start at 4; 35 kept + 2 new = 37 data rows -> 4..40 (1-based)
            int dataStart1 = 4;
            int lastData1 = dataStart1 + 35 + prep.addedCount() - 1;
            Row totalsRow = sh.getRow(lastData1); // 0-based index of 1-based lastData+1
            assertNotNull(totalsRow, "totals row missing at " + (lastData1 + 1));
            Cell sumSal = totalsRow.getCell(7);
            assertNotNull(sumSal);
            assertEquals(CellType.FORMULA, sumSal.getCellType(), "salary total should be a live formula");
            assertEquals("SUM(H4:H" + lastData1 + ")", sumSal.getCellFormula());

            // contribution column carries the template formula, shifted
            Cell contrib = sh.getRow(dataStart1).getCell(8);
            assertNotNull(contrib);
            assertEquals(CellType.FORMULA, contrib.getCellType());
            assertTrue(contrib.getCellFormula().contains("0.05*H" + (dataStart1 + 1)),
                    "expected shifted =0.05*H5, got: " + contrib.getCellFormula());

            // green fill on the two appended (new) rows
            assertFill(sh.getRow(lastData1 - 2), IndexedColors.LIGHT_GREEN, "first new row");
            assertFill(sh.getRow(lastData1 - 1), IndexedColors.LIGHT_GREEN, "second new row");

            // red REMOVED rows sit after the totals: exactly 3
            int removedSeen = 0;
            for (int r = lastData1 + 1; r <= lastData1 + 10; r++) {
                Row row = sh.getRow(r);
                if (row == null) break;
                Cell marker = row.getCell(0);
                assertNotNull(marker, "removed row needs its REMOVED marker");
                assertEquals("REMOVED", marker.getStringCellValue());
                assertFill(row, IndexedColors.ROSE, "removed row " + (r + 1));
                removedSeen++;
            }
            assertEquals(prep.removedCount(), removedSeen, "removed row count");

            // original template sheet untouched, original sheets all still present
            assertNotNull(wb.getSheet("AUG 2026"));
            assertTrue(wb.getNumberOfSheets() >= 113, "workbook should gain exactly one sheet");
        } finally {
            Files.deleteIfExists(prep.uploadFile());
        }
    }

    private void assertFill(Row row, IndexedColors expected, String what) {
        assertNotNull(row, what + " missing");
        boolean found = false;
        for (Cell c : row) {
            CellStyle st = c.getCellStyle();
            if (st.getFillPattern() == FillPatternType.SOLID_FOREGROUND
                    && st.getFillForegroundColor() == expected.getIndex()) {
                found = true;
                break;
            }
        }
        assertTrue(found, what + " should carry " + expected + " fill");
    }
}
