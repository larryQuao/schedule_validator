package com.schedulevalidator;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.parse.XlsxScheduleParser;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Stress diagnostic: runs the new-month merge against EVERY monthly sheet of the sample
 * with a small synthetic upload, so formula/layout problems in any era of the workbook
 * surface with the offending sheet and formula named.
 */
class SheetMergeAllSheetsTest {

    private static final Path SAMPLE = Paths.get("Samples", "ANNE MARIE CONTRIBUTION REPORT.xlsx");

    @Test
    void mergeSurvivesEveryMonthlySheet() throws Exception {
        assertTrue(Files.isRegularFile(SAMPLE), "sample workbook missing");
        List<String> failures = new ArrayList<>();
        int ok = 0;
        int index = 0;
        try (var wb = org.apache.poi.ss.usermodel.WorkbookFactory.create(SAMPLE.toFile())) {
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                String name = wb.getSheetName(s);
                index++;
                try {
                    new com.schedulevalidator.parse.XlsxScheduleParser();
                    var parser = new com.schedulevalidator.parse.XlsxScheduleParser();
                    var sheet = parser.parseSingleSheet(wb.getSheetAt(s));
                    if (sheet == null || sheet.entries().isEmpty()) continue;
                    List<ContributionEntry> upload = new ArrayList<>();
                    int kept = Math.min(5, sheet.entries().size());
                    for (int i = 0; i < kept; i++) {
                        ContributionEntry e = sheet.entries().get(i);
                        upload.add(new ContributionEntry(i + 1, i + 1, e.memberCode(), e.ssNumber(),
                                e.surname(), e.firstName(), e.otherNames(), "", e.basicSalary(),
                                e.contribution(), "", ""));
                    }
                    upload.add(new ContributionEntry(kept + 1, kept + 1, "", "Z00000000000" + (index % 10),
                            "TEST", "MEMBER" + index, "", "", 500.0, 25.0, "", ""));
                    ScheduleSheet upSheet = new ScheduleSheet("TEST MONTH " + index, null, null,
                            "TEST MONTH " + index, 3, upload, null, null, 0, List.of(), null);
                    Path outDir = Files.createTempDirectory("sv-stress-");
                    new com.schedulevalidator.convert.SheetMerger().build(SAMPLE, name,
                            upSheet, upload, outDir);
                    ok++;
                } catch (Throwable t) {
                    failures.add(name + " -> " + t.getMessage());
                }
            }
        }
        System.out.println("MERGE_OK=" + ok + " MERGE_FAILED=" + failures.size());
        failures.forEach(f -> System.out.println("FAIL: " + f));
        assertTrue(failures.isEmpty(), failures.size() + " of " + (ok + failures.size())
                + " sheets failed the merge; see console output");
    }
}
