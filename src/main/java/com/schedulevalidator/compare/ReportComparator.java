package com.schedulevalidator.compare;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.model.ValidationIssue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Compares a reference report against an uploaded one, month by month. Entries are joined on
 * normalized SS number (falling back to normalized full name when the SS number is blank).
 * Exact match by default: any missing/extra row or value difference is reported, with both
 * values in the message.
 */
public final class ReportComparator {

    public static final double AMOUNT_EPSILON = 0.011;

    public List<ValidationIssue> compare(ContributionReport reference, ContributionReport upload) {
        List<ValidationIssue> issues = new ArrayList<>();

        Map<String, ScheduleSheet> refByPeriod = sheetsByPeriod(reference);
        Map<String, ScheduleSheet> upByPeriod = sheetsByPeriod(upload);

        for (Map.Entry<String, ScheduleSheet> en : refByPeriod.entrySet()) {
            ScheduleSheet ref = en.getValue();
            ScheduleSheet up = upByPeriod.get(en.getKey());
            if (up == null) {
                issues.add(ValidationIssue.error("COMPARISON", ref.sheetName(), "-",
                        "SHEET_MISSING_IN_UPLOAD",
                        "Period '" + en.getKey() + "' (" + ref.sheetName() + ") has no counterpart in the uploaded file."));
                continue;
            }
            compareSheets(ref, up, issues);
        }
        for (Map.Entry<String, ScheduleSheet> en : upByPeriod.entrySet()) {
            if (!refByPeriod.containsKey(en.getKey())) {
                issues.add(ValidationIssue.warning("COMPARISON", en.getValue().sheetName(), "-",
                        "SHEET_NEW_IN_UPLOAD",
                        "Uploaded file contains period '" + en.getKey() + "' that is not in the reference."));
            }
        }
        return issues;
    }

    private void compareSheets(ScheduleSheet ref, ScheduleSheet up, List<ValidationIssue> out) {
        Map<String, ContributionEntry> refByKey = byKey(ref.entries());
        Map<String, ContributionEntry> upByKey = byKey(up.entries());

        for (Map.Entry<String, ContributionEntry> en : refByKey.entrySet()) {
            ContributionEntry u = upByKey.get(en.getKey());
            ContributionEntry r = en.getValue();
            if (u == null) {
                out.add(ValidationIssue.error("COMPARISON", ref.sheetName(), "ref row " + r.rowNumber(),
                        "ROW_MISSING_IN_UPLOAD",
                        r.displayName() + " (SS " + en.getKey() + ") is in the reference but missing from the upload."));
                continue;
            }
            if (r.basicSalary() != null && u.basicSalary() != null
                    && Math.abs(r.basicSalary() - u.basicSalary()) > AMOUNT_EPSILON) {
                out.add(ValidationIssue.error("COMPARISON", ref.sheetName(), "ref row " + r.rowNumber()
                                + " / upload row " + u.rowNumber(),
                        "SALARY_MISMATCH",
                        String.format("%s: basic salary %.2f in reference vs %.2f in upload.",
                                r.displayName(), r.basicSalary(), u.basicSalary())));
            }
            if (r.contribution() != null && u.contribution() != null
                    && Math.abs(r.contribution() - u.contribution()) > AMOUNT_EPSILON) {
                out.add(ValidationIssue.error("COMPARISON", ref.sheetName(), "ref row " + r.rowNumber()
                                + " / upload row " + u.rowNumber(),
                        "CONTRIBUTION_MISMATCH",
                        String.format("%s: contribution %.2f in reference vs %.2f in upload.",
                                r.displayName(), r.contribution(), u.contribution())));
            }
            if (!r.normalizedName().equals(u.normalizedName())) {
                out.add(ValidationIssue.warning("COMPARISON", ref.sheetName(),
                                "ref row " + r.rowNumber() + " / upload row " + u.rowNumber(),
                        "NAME_MISMATCH",
                        "Name differs: '" + r.displayName() + "' vs '" + u.displayName() + "'."));
            }
        }
        for (Map.Entry<String, ContributionEntry> en : upByKey.entrySet()) {
            if (!refByKey.containsKey(en.getKey())) {
                ContributionEntry u = en.getValue();
                out.add(ValidationIssue.warning("COMPARISON", up.sheetName(), "upload row " + u.rowNumber(),
                        "ROW_EXTRA_IN_UPLOAD",
                        u.displayName() + " (SS " + en.getKey() + ") appears in the upload but not in the reference."));
            }
        }

        if (ref.statedTotalContribution() != null && up.statedTotalContribution() != null
                && Math.abs(ref.statedTotalContribution() - up.statedTotalContribution()) > AMOUNT_EPSILON) {
            out.add(ValidationIssue.warning("COMPARISON", ref.sheetName(), "totals",
                    "STATED_TOTAL_DIFFERS",
                    String.format("Printed contribution total differs: %.2f (reference) vs %.2f (upload).",
                            ref.statedTotalContribution(), up.statedTotalContribution())));
        }
    }

    private Map<String, ScheduleSheet> sheetsByPeriod(ContributionReport report) {
        Map<String, ScheduleSheet> byPeriod = new HashMap<>();
        for (ScheduleSheet s : report.getMonthlySheets()) {
            String key = s.periodStart() != null ? s.periodStart() : s.sheetName().toUpperCase();
            byPeriod.putIfAbsent(key, s);
        }
        return byPeriod;
    }

    private Map<String, ContributionEntry> byKey(List<ContributionEntry> entries) {
        Map<String, ContributionEntry> map = new HashMap<>();
        Set<String> usedNames = new HashSet<>();
        for (ContributionEntry e : entries) {
            String key = e.normalizedSs();
            if (key.isEmpty()) {
                key = "NAME:" + e.normalizedName();
                if (usedNames.contains(key) || key.equals("NAME:")) continue;
            }
            usedNames.add(key);
            map.putIfAbsent(key, e);
        }
        return map;
    }
}
