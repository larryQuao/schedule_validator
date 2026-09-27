package com.schedulevalidator;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ScheduleSheet;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

/** Reproduces the "named H does not exist" failure against non-monthly sheets. */
class NonMonthlySheetsTest {

    private static final Path SAMPLE = Paths.get("Samples", "ANNE MARIE CONTRIBUTION REPORT.xlsx");

    @Test
    void mergeAgainstNonMonthlySheets() throws Exception {
        assertTrue(Files.isRegularFile(SAMPLE), "sample workbook missing");
        String[] names = {"DETAILS", "SUMMARY", "SUSPENSE ACCOUNT", "SUSPENSE-27-01-2020",
                "SUSPENSE-30-03-2021", "SUSPENSE-02-08-2021", "SUSPENSE-18-02-2022",
                "SURCHARGE", "Sheet1", "Sheet5"};
        List<String> failures = new ArrayList<>();
        for (String name : names) {
            try {
                var parser = new com.schedulevalidator.parse.XlsxScheduleParser();
                ScheduleSheet sheet;
                try (var wb = org.apache.poi.ss.usermodel.WorkbookFactory.create(
                        SAMPLE.toFile(), null, true)) {
                    sheet = parser.parseSingleSheet(wb.getSheet(name));
                }
                if (sheet == null || sheet.entries().isEmpty()) {
                    System.out.println("SKIP (no schedule): " + name);
                    continue;
                }
                List<ContributionEntry> upload = new ArrayList<>();
                int kept = Math.min(5, sheet.entries().size());
                for (int i = 0; i < kept; i++) {
                    ContributionEntry e = sheet.entries().get(i);
                    upload.add(new ContributionEntry(i + 1, i + 1, e.memberCode(), e.ssNumber(),
                            e.surname(), e.firstName(), e.otherNames(), "", e.basicSalary(),
                            e.contribution(), "", ""));
                }
                upload.add(new ContributionEntry(kept + 1, kept + 1, "", "Z000000000009",
                        "TEST", "MEMBER", "", "", 500.0, 25.0, "", ""));
                ScheduleSheet upSheet = new ScheduleSheet("TEST MONTH", null, null,
                        "TEST MONTH", 3, upload, null, null, 0, List.of(), null);
                Path outDir = Files.createTempDirectory("sv-nonmonthly-");
                var summary = new com.schedulevalidator.convert.SheetMerger().build(
                        SAMPLE, name, upSheet, upload, outDir);
                System.out.println("OK: " + name + " (matched=" + summary.matched()
                        + " added=" + summary.added() + " removed=" + summary.removed() + ")");
            } catch (Throwable t) {
                failures.add(name + " -> " + t.getMessage());
            }
        }
        failures.forEach(f -> System.out.println("FAIL: " + f));
        assertTrue(failures.isEmpty(), String.join(" | ", failures));
    }
}
