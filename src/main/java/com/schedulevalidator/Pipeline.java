package com.schedulevalidator;

import com.schedulevalidator.compare.ReportComparator;
import com.schedulevalidator.convert.SheetMerger;
import com.schedulevalidator.convert.XlsxExporter;
import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import com.schedulevalidator.model.ValidationIssue;
import com.schedulevalidator.parse.ImageScheduleParser;
import com.schedulevalidator.parse.PdfScheduleParser;
import com.schedulevalidator.parse.XlsxScheduleParser;
import com.schedulevalidator.validate.ValidationEngine;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/** Shared pipeline: parse (by format) -> validate -> compare -> export. Used by GUI and CLI. */
public final class Pipeline {

    public record Result(ContributionReport uploadReport,
                         ContributionReport referenceReport,
                         List<ValidationIssue> issues,
                         Path reportFile,
                         Path convertedFile,
                         Path updatedWorkbook,
                         Path finalReportFile,
                         List<SheetMerger.RecordRow> records) {}

    private Pipeline() {}

    /** tessdata lookup: TESSDATA_PREFIX env, then working dir, then beside the app jars. */
    static Path resolveTessdataDir() {
        String env = System.getenv("TESSDATA_PREFIX");
        if (env != null && !env.isBlank()) return Path.of(env);
        Path cwd = Path.of("tessdata");
        if (Files.isDirectory(cwd)) return cwd;
        try {
            Path appDir = Path.of(Pipeline.class.getProtectionDomain()
                    .getCodeSource().getLocation().toURI()).getParent();
            Path beside = appDir.resolve("tessdata");
            if (Files.isDirectory(beside)) return beside;
        } catch (Exception ignored) {
        }
        return cwd;
    }

    /** Writable output location: user Documents (install dir may be read-only). */
    static Path defaultOutDir() {
        String home = System.getProperty("user.home");
        if (home != null && !home.isBlank()) {
            return Path.of(home, "Documents", "Schedule Validator", "out");
        }
        return Path.of("out");
    }

    public static ContributionReport parse(Path file, Path tessdataDir) throws IOException {
        String ext = extension(file);
        return switch (ext) {
            case "xlsx", "xlsm", "xls" -> new XlsxScheduleParser().parse(file);
            case "pdf" -> new PdfScheduleParser().parse(file.toFile());
            case "png", "jpg", "jpeg", "tif", "tiff", "bmp", "gif" ->
                    new ImageScheduleParser(tessdataDir).parse(file);
            default -> throw new IOException("Unsupported file type ." + ext
                    + " - use .xlsx, .pdf or an image (png/jpg/tiff/bmp).");
        };
    }

    public record Outputs(Path reportFile, Path convertedFile) {}

    public static Result run(Path referencePath, Path uploadPath, Path outDir,
                             Path tessdataDir, List<String> log) throws IOException {
        ValidationEngine engine = new ValidationEngine();
        ReportComparator comparator = new ReportComparator();
        XlsxExporter exporter = new XlsxExporter();
        List<ValidationIssue> issues = new ArrayList<>();

        ContributionReport refReport = null;
        ContributionReport upReport = null;

        if (uploadPath != null) {
            log.add("Parsing uploaded file: " + uploadPath.getFileName());
            upReport = parse(uploadPath, tessdataDir);
            issues.addAll(engine.validate(upReport));
        }
        if (referencePath != null) {
            log.add("Parsing reference file: " + referencePath.getFileName());
            refReport = parse(referencePath, tessdataDir);
            if (upReport == null) {
                issues.addAll(engine.validate(refReport));
            } else {
                log.add("Comparing upload against reference...");
                issues.addAll(comparator.compare(refReport, upReport));
            }
        }
        if (upReport == null && refReport == null) {
            throw new IOException("Nothing to validate: pick a reference and/or an uploaded file.");
        }

        Files.createDirectories(outDir);
        Path reportFile = outDir.resolve("validation_report.xlsx");
        exporter.writeValidationReport(issues, referencePath, uploadPath, reportFile);
        log.add("Wrote " + reportFile);

        Path convertedFile = null;
        boolean needsConversion = upReport != null && upReport.getKind() != ContributionReport.SourceKind.XLSX;
        if (needsConversion) {
            convertedFile = outDir.resolve("converted_schedule.xlsx");
            exporter.writeConvertedSchedule(upReport, convertedFile);
            log.add("Wrote " + convertedFile + " (converted to .xlsx in reference layout)");
        }
        return new Result(upReport, refReport, issues, reportFile, convertedFile, null, null, null);
    }

    /**
     * Sheet-merge mode: the user picks one sheet of the reference report, the upload carries
     * the new month's schedule, and after validation a NEW sheet is added to a copy of the
     * reference workbook — template structure and formulas preserved, new records green,
     * records missing from the upload marked REMOVED below the totals in red.
     */
    public static Result runSheetMerge(Path referencePath, String sheetName, Path uploadPath,
                                       Path outDir, Path tessdataDir, List<String> log) throws IOException {
        log.add("Parsing uploaded file: " + uploadPath.getFileName());
        ContributionReport upReport = parse(uploadPath, tessdataDir);
        List<ValidationIssue> issues = new ArrayList<>(new ValidationEngine().validate(upReport));

        List<ScheduleSheet> uploads = upReport.getMonthlySheets();
        if (uploads.isEmpty()) {
            throw new IOException("Uploaded file contains no recognizable schedule rows.");
        }
        List<ContributionEntry> uploadEntries = new java.util.ArrayList<>();
        uploads.forEach(s -> uploadEntries.addAll(s.entries()));
        // name the new sheet after the NEWEST month in the upload
        ScheduleSheet newestUpload = uploads.get(0);
        for (ScheduleSheet s : uploads) {
            if (s.periodStart() != null && (newestUpload.periodStart() == null
                    || s.periodStart().compareTo(newestUpload.periodStart()) > 0)) {
                newestUpload = s;
            }
        }

        log.add("Parsing reference workbook: " + referencePath.getFileName());
        ContributionReport refReport = parse(referencePath, tessdataDir);
        ScheduleSheet selected = refReport.getMonthlySheets().stream()
                .filter(s -> s.sheetName().equalsIgnoreCase(sheetName))
                .findFirst()
                .orElseThrow(() -> new IOException(
                        "Sheet '" + sheetName + "' was not found or has no recognizable schedule in the reference."));

        ContributionReport selectedOnly = new ContributionReport(referencePath, ContributionReport.SourceKind.XLSX);
        selectedOnly.getMonthlySheets().add(selected);
        issues.addAll(new ReportComparator().compare(selectedOnly, upReport));

        Files.createDirectories(outDir);
        log.add("Building new month sheet from template '" + selected.sheetName() + "'...");
        SheetMerger.MergeSummary merge = new SheetMerger().build(
                referencePath, selected.sheetName(), newestUpload, uploadEntries, outDir);
        issues.add(ValidationIssue.info("MERGE", merge.newSheetName(), "-", "MERGE_SUMMARY",
                merge.matched() + " carried over, " + merge.added() + " new (green), "
                        + merge.removed() + " removed (red). Saved: " + merge.updatedWorkbook()));
        for (String note : merge.notes()) {
            issues.add(ValidationIssue.warning("MERGE", merge.newSheetName(), "-", "MERGE_NOTE", note));
        }
        log.add("Wrote " + merge.updatedWorkbook() + " (new sheet: " + merge.newSheetName() + ")");

        Path reportFile = outDir.resolve("validation_report.xlsx");
        new XlsxExporter().writeValidationReport(issues, referencePath, uploadPath, reportFile);
        log.add("Wrote " + reportFile);

        Path finalReport = outDir.resolve("final_validation_report.xlsx");
        new XlsxExporter().writeFinalValidationReport(merge.records(), issues, referencePath,
                selected.sheetName(), uploadPath, merge.newSheetName(),
                merge.matched(), merge.added(), merge.removed(), finalReport);
        log.add("Wrote " + finalReport + " (summary + final records + findings)");
        return new Result(upReport, refReport, issues, reportFile, null, merge.updatedWorkbook(),
                finalReport, merge.records());
    }

    static String extension(Path file) {
        String name = file.getFileName().toString().toLowerCase(Locale.ROOT);
        int dot = name.lastIndexOf('.');
        return dot < 0 ? "" : name.substring(dot + 1);
    }
}
