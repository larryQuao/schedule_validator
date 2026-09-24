package com.schedulevalidator.parse;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.Word;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

/**
 * OCR pipeline for image uploads (png/jpg/tiff/...): Tesseract word boxes are clustered into
 * visual lines, then {@link TableExtractor} reconstructs the schedule columns. Requires the
 * tessdata directory (eng.traineddata) shipped next to the app. Accuracy depends on scan
 * quality; every converted row keeps its source line number for manual checking.
 */
public final class ImageScheduleParser {

    private final Path tessdataDir;

    public ImageScheduleParser(Path tessdataDir) {
        this.tessdataDir = tessdataDir;
    }

    public ContributionReport parse(Path image) throws IOException {
        ContributionReport report = new ContributionReport(image, ContributionReport.SourceKind.IMAGE);
        if (!Files.isRegularFile(tessdataDir.resolve("eng.traineddata"))) {
            report.addParseNote("tessdata/eng.traineddata not found next to the app - OCR unavailable.");
            return report;
        }
        BufferedImage img = ImageIO.read(image.toFile());
        if (img == null) {
            report.addParseNote("Not a readable image file: " + image.getFileName());
            return report;
        }

        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath(tessdataDir.toAbsolutePath().toString());
        tesseract.setLanguage("eng");
        tesseract.setPageSegMode(6); // assume a uniform block of text
        List<Word> words;
        try {
            words = tesseract.getWords(img, 1); // 1 = RIL_WORD
        } catch (Exception e) {
            report.addParseNote("OCR failed: " + e.getMessage());
            return report;
        }

        // tess4j Word bounding boxes: x, y, width, height
        record BoxWord(String text, float x0, float x1, float yCenter) {}
        List<BoxWord> boxed = new ArrayList<>();
        for (Word w : words) {
            String text = w.getText();
            if (text == null || text.isBlank()) continue;
            var b = w.getBoundingBox();
            boxed.add(new BoxWord(text.trim(), b.x, b.x + b.width, b.y + b.height / 2f));
        }
        boxed.sort(Comparator.comparingDouble(BoxWord::yCenter).thenComparingDouble(BoxWord::x0));

        // cluster into lines by y-center proximity
        List<List<TableExtractor.Word>> lines = new ArrayList<>();
        List<TableExtractor.Word> current = new ArrayList<>();
        float currentY = Float.NaN;
        for (BoxWord w : boxed) {
            if (!Float.isNaN(currentY) && Math.abs(w.yCenter - currentY) > 12f) {
                lines.add(current);
                current = new ArrayList<>();
            }
            currentY = w.yCenter;
            current.add(new TableExtractor.Word(w.text, w.x0, w.x1));
        }
        if (!current.isEmpty()) lines.add(current);

        TableExtractor.Table table = new TableExtractor().extractFromLines(lines);
        if (!table.found()) {
            report.addParseNote("OCR: no schedule header row detected in image - "
                    + lines.size() + " text lines captured but not recognised as a schedule.");
            return report;
        }

        // map headers to canonical roles
        int ss = -1, surname = -1, first = -1, other = -1, salary = -1, contrib = -1, member = -1;
        for (int c = 0; c < table.headers.size(); c++) {
            String h = table.headers.get(c);
            if (h.contains("MEMBER")) member = member < 0 ? c : member;
            else if (h.contains("SURNAME")) surname = surname < 0 ? c : surname;
            else if (h.contains("FIRST")) first = first < 0 ? c : first;
            else if (h.contains("OTHER")) other = other < 0 ? c : other;
            else if (h.contains("SALARY")) salary = salary < 0 ? c : salary;
            else if (h.contains("CONTRIBUTION")) contrib = contrib < 0 ? c : contrib;
            else if (h.startsWith("SS") || h.startsWith("SSNIT")) ss = ss < 0 ? c : ss;
        }
        if (ss < 0 && surname < 0) {
            report.addParseNote("OCR: header found but columns could not be mapped confidently.");
            return report;
        }

        List<ContributionEntry> entries = new ArrayList<>();
        int rowNo = 0;
        for (List<String> row : table.rows) {
            rowNo++;
            String ssNo = cell(row, ss);
            String sur = cell(row, surname);
            Double sal = num(cell(row, salary));
            Double con = num(cell(row, contrib));
            boolean hasIdentity = !ssNo.isBlank() || !sur.isBlank();
            if (!hasIdentity || (sal == null && con == null)) continue;
            entries.add(new ContributionEntry(rowNo, parseSeq(cell(row, 0)), cell(row, member),
                    ssNo, sur, cell(row, first), cell(row, other), "", sal, con, "", ""));
        }
        if (entries.isEmpty()) {
            report.addParseNote("OCR: header detected but no member rows could be read.");
            return report;
        }
        String pageName = image.getFileName().toString().toUpperCase(Locale.ROOT);
        report.getMonthlySheets().add(new ScheduleSheet(pageName, null, null, pageName,
                table.headerLineIndex + 1, entries, null, null, 0, new ArrayList<>(), null));
        return report;
    }

    private static String cell(List<String> row, int idx) {
        return idx >= 0 && idx < row.size() ? row.get(idx).trim() : "";
    }

    private static Integer parseSeq(String s) {
        try { return Integer.parseInt(s.trim()); } catch (NumberFormatException e) { return null; }
    }

    private static Double num(String s) {
        if (s == null) return null;
        String t = s.replace(",", "").replace("O", "0").replace("o", "0").trim(); // OCR digit fixes
        if (t.isEmpty()) return null;
        try { return Double.parseDouble(t); } catch (NumberFormatException e) { return null; }
    }
}
