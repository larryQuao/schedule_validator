package com.schedulevalidator.model;

import java.util.List;

/**
 * One monthly schedule table (one sheet in the reference workbook, or one detected table in a
 * PDF/image upload).
 *
 * @param sheetName              source sheet name (or "PAGE n" for pdf/image)
 * @param periodStart            parsed period start as {@code yyyy-MM}, null if not parseable
 * @param periodEnd              parsed period end as {@code yyyy-MM} (range sheets), null if single month
 * @param periodLabel            original period text (sheet name)
 * @param headerRow              1-based row number the column headers were found on
 * @param entries                member rows, in source order
 * @param statedTotalSalary      salary total printed on the sheet, null if none
 * @param statedTotalContribution contribution total printed on the sheet, null if none
 * @param totalRow               1-based row number of the totals row, 0 if none
 * @param postTotalRows          row numbers of member rows found BELOW the totals row
 * @param monthLabelInSheet      month written inside the sheet (e.g. "MONTH:SEPTEMBER"), null if none
 */
public record ScheduleSheet(
        String sheetName,
        String periodStart,
        String periodEnd,
        String periodLabel,
        int headerRow,
        List<ContributionEntry> entries,
        Double statedTotalSalary,
        Double statedTotalContribution,
        int totalRow,
        List<Integer> postTotalRows,
        String monthLabelInSheet) {

    public double computedTotalSalary() {
        return entries.stream().mapToDouble(e -> e.basicSalary() == null ? 0 : e.basicSalary()).sum();
    }

    public double computedTotalContribution() {
        return entries.stream().mapToDouble(e -> e.contribution() == null ? 0 : e.contribution()).sum();
    }
}
