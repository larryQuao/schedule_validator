package com.schedulevalidator;

import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.model.ValidationIssue;
import com.schedulevalidator.parse.XlsxScheduleParser;
import com.schedulevalidator.validate.ValidationEngine;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Integration test against the shipped sample workbook. Skips quietly when the sample is not
 * present so the build still works on machines without it.
 */
class SampleParseTest {

    private static final Path SAMPLE = Paths.get(
            "Samples", "ANNE MARIE CONTRIBUTION REPORT.xlsx");

    private ContributionReport sample() throws Exception {
        assertTrue(Files.isRegularFile(SAMPLE), "sample workbook missing: " + SAMPLE.toAbsolutePath());
        return new XlsxScheduleParser().parse(SAMPLE);
    }

    @Test
    void parsesManyMonthlySheets() throws Exception {
        ContributionReport report = sample();
        assertTrue(report.getMonthlySheets().size() >= 100,
                "expected ~106 monthly sheets, got " + report.getMonthlySheets().size());
        assertNotNull(report.getEmployerName());
        assertTrue(report.getEmployerName().toUpperCase().contains("ANN"), "employer: " + report.getEmployerName());
    }

    @Test
    void parsesAugust2026WithKnownEntries() throws Exception {
        ContributionReport report = sample();
        Optional<ScheduleSheet> aug26 = report.getMonthlySheets().stream()
                .filter(s -> s.periodStart() != null && s.periodStart().equals("2026-08"))
                .findFirst();
        assertTrue(aug26.isPresent(), "AUG 2026 sheet not found/parsed");
        ScheduleSheet sheet = aug26.get();
        assertEquals(38, sheet.entries().size(), "AUG 2026 should have 33 + 5 appended members");
        assertNotNull(sheet.statedTotalContribution(), "totals row not detected");
        assertEquals(1290.0, sheet.statedTotalContribution(), 0.01, "printed contribution total");
        assertEquals(1460.0, sheet.computedTotalContribution(), 0.01,
                "33 members x their 5% + 5 appended members = 1290 + 170");
    }

    @Test
    void engineFlagsTotalMismatchOnAugust2026() throws Exception {
        ContributionReport report = sample();
        List<ValidationIssue> issues = new ValidationEngine().validate(report);
        assertTrue(issues.stream().anyMatch(i ->
                        "TOTAL_CONTRIBUTION_MISMATCH".equals(i.rule())
                                && i.sheet().toUpperCase().contains("AUG 2026")),
                "expected total-mismatch finding for AUG 2026");
        assertTrue(issues.stream().anyMatch(i -> "ENTRIES_BELOW_TOTALS".equals(i.rule())),
                "expected entries-below-totals finding");
    }

    @Test
    void engineFlagsMonthLabelMismatch() throws Exception {
        ContributionReport report = sample();
        List<ValidationIssue> issues = new ValidationEngine().validate(report);
        assertTrue(issues.stream().anyMatch(i ->
                        "MONTH_LABEL_MISMATCH".equals(i.rule())
                                && i.sheet().toUpperCase().contains("AUGUST 2017")),
                "expected the AUGUST 2017 sheet (labelled SEPTEMBER inside) to be flagged");
    }
}
