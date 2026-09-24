package com.schedulevalidator;

import com.schedulevalidator.compare.ReportComparator;
import com.schedulevalidator.convert.XlsxExporter;
import com.schedulevalidator.model.ContributionReport;
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
                         Path convertedFile) {}

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
        return new Result(upReport, refReport, issues, reportFile, convertedFile);
    }

    static String extension(Path file) {
        String name = file.getFileName().toString().toLowerCase(Locale.ROOT);
        int dot = name.lastIndexOf('.');
        return dot < 0 ? "" : name.substring(dot + 1);
    }
}
