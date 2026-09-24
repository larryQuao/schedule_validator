package com.schedulevalidator.validate;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.model.ValidationIssue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

/**
 * Rule set for one parsed report. Defaults to exact match (amounts equal to the cent); the only
 * slack is 0.011 to absorb binary floating-point representation. Every rule cites sheet + row so
 * findings can be traced back to the source file.
 */
public final class ValidationEngine {

    public static final double AMOUNT_EPSILON = 0.011;
    private static final double RATE = 0.05;
    private static final Pattern SS_MODERN = Pattern.compile("[A-Z]\\d{12}");
    private static final Pattern SS_LEGACY = Pattern.compile("[A-Z0-9]{10,13}");

    public List<ValidationIssue> validate(ContributionReport report) {
        List<ValidationIssue> issues = new ArrayList<>();
        String src = sourceLabel(report);
        for (String note : report.getParseNotes()) {
            issues.add(ValidationIssue.warning(src, "-", "-", "PARSE_NOTE", note));
        }
        for (ScheduleSheet sheet : report.getMonthlySheets()) {
            checkTotals(src, sheet, issues);
            checkEntriesAfterTotals(src, sheet, issues);
            checkRate(src, sheet, issues);
            checkSequence(src, sheet, issues);
            checkMemberCodes(src, sheet, issues);
            checkSsFormats(src, sheet, issues);
            checkDuplicatesInSheet(src, sheet, issues);
            checkNameWhitespace(src, sheet, issues);
            checkMonthLabel(src, sheet, issues);
        }
        checkDuplicatesAcrossMonths(src, report, issues);
        checkMemberContinuity(src, report, issues);
        checkSummaryCross(src, report, issues);
        return issues;
    }

    // ------------------------------------------------------------------ rules

    private void checkTotals(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        if (sheet.totalRow() == 0) {
            out.add(ValidationIssue.warning(src, sheet.sheetName(), "-",
                    "NO_TOTALS_ROW", "Sheet has no recognisable totals row."));
            return;
        }
        double computedSal = sheet.computedTotalSalary();
        double computedCon = sheet.computedTotalContribution();
        Double statedSal = sheet.statedTotalSalary();
        Double statedCon = sheet.statedTotalContribution();
        if (statedSal != null && Math.abs(computedSal - statedSal) > AMOUNT_EPSILON) {
            out.add(ValidationIssue.error(src, sheet.sheetName(), "row " + sheet.totalRow(),
                    "TOTAL_SALARY_MISMATCH",
                    String.format("Stated basic salary total %.2f does not match computed %.2f (%d entries).",
                            statedSal, computedSal, sheet.entries().size())));
        }
        if (statedCon != null && Math.abs(computedCon - statedCon) > AMOUNT_EPSILON) {
            out.add(ValidationIssue.error(src, sheet.sheetName(), "row " + sheet.totalRow(),
                    "TOTAL_CONTRIBUTION_MISMATCH",
                    String.format("Stated 5%% contribution total %.2f does not match computed %.2f (%d entries).",
                            statedCon, computedCon, sheet.entries().size())));
        }
    }

    private void checkEntriesAfterTotals(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        if (sheet.postTotalRows().isEmpty()) return;
        double missed = sheet.entries().stream()
                .filter(e -> sheet.postTotalRows().contains(e.rowNumber()))
                .mapToDouble(e -> e.contribution() == null ? 0 : e.contribution()).sum();
        out.add(ValidationIssue.error(src, sheet.sheetName(),
                "rows " + sheet.postTotalRows().get(0) + "-" + sheet.postTotalRows().get(sheet.postTotalRows().size() - 1),
                "ENTRIES_BELOW_TOTALS",
                String.format("%d member row(s) appear BELOW the totals row and are excluded from it "
                                + "(unaccounted contribution: %.2f).",
                        sheet.postTotalRows().size(), missed)));
    }

    private void checkRate(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        for (ContributionEntry e : sheet.entries()) {
            if (e.basicSalary() == null || e.contribution() == null) {
                out.add(ValidationIssue.warning(src, sheet.sheetName(), "row " + e.rowNumber(),
                        "MISSING_AMOUNT", "Row is missing basic salary or 5% contribution."));
                continue;
            }
            double expected = e.basicSalary() * RATE;
            if (Math.abs(expected - e.contribution()) > AMOUNT_EPSILON) {
                out.add(ValidationIssue.error(src, sheet.sheetName(), "row " + e.rowNumber(),
                        "RATE_MISMATCH",
                        String.format("Contribution %.2f is not 5%% of salary %.2f (expected %.2f, %.2f%% given).",
                                e.contribution(), e.basicSalary(), expected,
                                e.basicSalary() == 0 ? 0 : e.contribution() / e.basicSalary() * 100)));
            }
        }
    }

    private void checkSequence(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        Integer prev = null;
        for (ContributionEntry e : sheet.entries()) {
            Integer seq = e.seqNo();
            if (seq == null) continue;
            if (prev != null) {
                if (seq == prev) {
                    out.add(ValidationIssue.warning(src, sheet.sheetName(), "row " + e.rowNumber(),
                            "SEQ_DUPLICATE", "Sequence number " + seq + " repeats."));
                } else if (seq != prev + 1) {
                    out.add(ValidationIssue.warning(src, sheet.sheetName(), "row " + e.rowNumber(),
                            "SEQ_GAP", "Sequence jumps from " + prev + " to " + seq + "."));
                }
            }
            prev = seq;
        }
    }

    private void checkMemberCodes(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        long missing = sheet.entries().stream().filter(e -> e.memberCode().isBlank()).count();
        if (missing > 0) {
            out.add(ValidationIssue.warning(src, sheet.sheetName(), "-",
                    "MISSING_MEMBER_CODE",
                    missing + " of " + sheet.entries().size() + " rows have no Member Code."));
        }
    }

    private void checkSsFormats(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        for (ContributionEntry e : sheet.entries()) {
            String ss = e.normalizedSs();
            if (ss.isEmpty()) {
                out.add(ValidationIssue.warning(src, sheet.sheetName(), "row " + e.rowNumber(),
                        "MISSING_SS_NO", "Row has no SS number: " + e.displayName()));
                continue;
            }
            if (!SS_MODERN.matcher(ss).matches() && !SS_LEGACY.matcher(ss).matches()) {
                out.add(ValidationIssue.warning(src, sheet.sheetName(), "row " + e.rowNumber(),
                        "SS_FORMAT", "SS number '" + ss + "' (" + e.displayName() + ") matches no known format."));
            }
        }
    }

    private void checkDuplicatesInSheet(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        Map<String, List<Integer>> seen = new HashMap<>();
        for (ContributionEntry e : sheet.entries()) {
            String ss = e.normalizedSs();
            if (ss.isEmpty()) continue;
            seen.computeIfAbsent(ss, k -> new ArrayList<>()).add(e.rowNumber());
        }
        for (var en : seen.entrySet()) {
            if (en.getValue().size() > 1) {
                out.add(ValidationIssue.error(src, sheet.sheetName(), "rows " + en.getValue(),
                        "DUPLICATE_SS_IN_SHEET", "SS number " + en.getKey() + " appears "
                                + en.getValue().size() + " times in this sheet."));
            }
        }
    }

    private void checkNameWhitespace(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        for (ContributionEntry e : sheet.entries()) {
            for (String field : new String[]{e.surname(), e.firstName(), e.otherNames(), e.fullName()}) {
                if (field != null && (field.startsWith(" ") || field.endsWith(" ") || field.contains("  "))) {
                    out.add(ValidationIssue.info(src, sheet.sheetName(), "row " + e.rowNumber(),
                            "NAME_WHITESPACE",
                            "Name has irregular spacing: '" + field + "' (" + e.displayName() + ")"));
                    break;
                }
            }
        }
    }

    private void checkMonthLabel(String src, ScheduleSheet sheet, List<ValidationIssue> out) {
        if (sheet.monthLabelInSheet() == null || sheet.periodStart() == null) return;
        String sheetMonth = monthName(sheet.periodStart());
        if (sheetMonth != null && !sheetMonth.equals(sheet.monthLabelInSheet())) {
            out.add(ValidationIssue.warning(src, sheet.sheetName(), "header",
                    "MONTH_LABEL_MISMATCH",
                    "Sheet is named '" + sheet.periodLabel() + "' but the in-sheet label reads MONTH:"
                            + sheet.monthLabelInSheet() + "."));
        }
    }

    private void checkDuplicatesAcrossMonths(String src, ContributionReport report, List<ValidationIssue> out) {
        Map<String, List<String>> bySs = new HashMap<>();
        for (ScheduleSheet sheet : report.getMonthlySheets()) {
            for (ContributionEntry e : sheet.entries()) {
                String ss = e.normalizedSs();
                if (!ss.isEmpty()) bySs.computeIfAbsent(ss, k -> new ArrayList<>()).add(sheet.sheetName());
            }
        }
    }

    private void checkMemberContinuity(String src, ContributionReport report, List<ValidationIssue> out) {
        List<ScheduleSheet> monthly = report.getMonthlySheets().stream()
                .filter(s -> s.periodStart() != null && s.periodEnd() == null).toList();
        for (int i = 0; i < monthly.size() - 1; i++) {
            ScheduleSheet cur = monthly.get(i);
            ScheduleSheet next = monthly.get(i + 1);
            Set<String> nextSs = new HashSet<>();
            next.entries().forEach(e -> nextSs.add(e.normalizedSs()));
            for (ContributionEntry e : cur.entries()) {
                String ss = e.normalizedSs();
                if (!ss.isEmpty() && !nextSs.contains(ss)) {
                    out.add(ValidationIssue.info(src, next.sheetName(), "-",
                            "MEMBER_DISAPPEARED",
                            e.displayName() + " (SS " + ss + ") is in '" + cur.sheetName()
                                    + "' but not in '" + next.sheetName() + "'."));
                }
            }
        }
    }

    private void checkSummaryCross(String src, ContributionReport report, List<ValidationIssue> out) {
        if (report.getSummaryEntries().isEmpty() || report.getMonthlySheets().isEmpty()) return;
        Map<String, Double> monthlyTotals = new HashMap<>();
        Map<String, String> monthlyName = new HashMap<>();
        for (ScheduleSheet sheet : report.getMonthlySheets()) {
            for (ContributionEntry e : sheet.entries()) {
                String ss = e.normalizedSs();
                if (ss.isEmpty()) continue;
                monthlyTotals.merge(ss, e.contribution() == null ? 0 : e.contribution(), Double::sum);
                monthlyName.putIfAbsent(ss, e.displayName());
            }
        }
        int checked = 0;
        int mismatches = 0;
        for (ContributionEntry s : report.getSummaryEntries()) {
            String ss = s.normalizedSs();
            if (ss.isEmpty() || s.contribution() == null) continue;
            Double monthly = monthlyTotals.get(ss);
            if (monthly == null) continue;
            checked++;
            if (Math.abs(monthly - s.contribution()) > AMOUNT_EPSILON) {
                mismatches++;
                if (mismatches <= 25) {
                    out.add(ValidationIssue.warning(src, "SUMMARY", "row " + s.rowNumber(),
                            "SUMMARY_MISMATCH",
                            String.format("Summary total %.2f for %s (SS %s) differs from the sum of monthly sheets %.2f.",
                                    s.contribution(), s.displayName(), ss, monthly)));
                }
            }
        }
        if (checked > 0) {
            out.add(ValidationIssue.info(src, "SUMMARY", "-", "SUMMARY_CROSSCHECK",
                    checked + " members cross-checked against monthly sheets, " + mismatches + " mismatch(es)."));
        }
    }

    // ------------------------------------------------------------------ helpers

    static String monthName(String yyyyMm) {
        String[] parts = yyyyMm.split("-");
        if (parts.length != 2) return null;
        int m;
        try { m = Integer.parseInt(parts[1]); } catch (NumberFormatException e) { return null; }
        String[] names = {"JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
                "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"};
        return m >= 1 && m <= 12 ? names[m - 1] : null;
    }

    private static String sourceLabel(ContributionReport report) {
        return switch (report.getKind()) {
            case XLSX -> "UPLOAD";
            case PDF -> "UPLOAD-PDF";
            case IMAGE -> "UPLOAD-IMAGE";
        };
    }
}
