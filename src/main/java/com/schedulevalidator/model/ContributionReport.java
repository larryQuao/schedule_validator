package com.schedulevalidator.model;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/** A whole parsed report: header info + monthly schedules + auxiliary sheets. */
public class ContributionReport {

    public enum SourceKind { XLSX, PDF, IMAGE }

    private final Path sourceFile;
    private final SourceKind kind;
    private String employerName;
    private String ssnitRegNumber;
    private String schemeName;
    private final List<ScheduleSheet> monthlySheets = new ArrayList<>();
    /** Flattened member rows from SUMMARY / DETAILS auxiliary sheets, if present. */
    private final List<ContributionEntry> summaryEntries = new ArrayList<>();
    private final List<SuspenseRecord> suspenseRecords = new ArrayList<>();
    /** Non-fatal notes about the parse itself (e.g. sheets that could not be interpreted). */
    private final List<String> parseNotes = new ArrayList<>();

    public ContributionReport(Path sourceFile, SourceKind kind) {
        this.sourceFile = sourceFile;
        this.kind = kind;
    }

    public Path getSourceFile() { return sourceFile; }
    public SourceKind getKind() { return kind; }

    public String getEmployerName() { return employerName; }
    public void setEmployerName(String employerName) { this.employerName = employerName; }
    public String getSsnitRegNumber() { return ssnitRegNumber; }
    public void setSsnitRegNumber(String ssnitRegNumber) { this.ssnitRegNumber = ssnitRegNumber; }
    public String getSchemeName() { return schemeName; }
    public void setSchemeName(String schemeName) { this.schemeName = schemeName; }

    public List<ScheduleSheet> getMonthlySheets() { return monthlySheets; }
    public List<ContributionEntry> getSummaryEntries() { return summaryEntries; }
    public List<SuspenseRecord> getSuspenseRecords() { return suspenseRecords; }
    public List<String> getParseNotes() { return parseNotes; }

    public void addParseNote(String note) { parseNotes.add(note); }

    public int totalEntries() {
        return monthlySheets.stream().mapToInt(s -> s.entries().size()).sum();
    }

    /** Record of an unallocated/excess payment on a SUSPENSE sheet. */
    public record SuspenseRecord(int rowNumber, String payer, Double amount, String description) {}
}
