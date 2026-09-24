package com.schedulevalidator.parse;

import com.schedulevalidator.model.ContributionEntry;
import com.schedulevalidator.model.ContributionReport;
import com.schedulevalidator.model.ScheduleSheet;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Extracts schedule tables from (text-based) PDFs. Words are collected with their x-positions
 * from PDFBox text positions, grouped into visual lines by y, then handed to
 * {@link TableExtractor} for column reconstruction. Each page becomes a candidate sheet.
 * Scanned PDFs (no text layer) come back with a parse note — route those through the image
 * pipeline instead.
 */
public final class PdfScheduleParser {

    private static final float LINE_TOLERANCE = 3.0f;

    public ContributionReport parse(File pdf) throws IOException {
        ContributionReport report = new ContributionReport(pdf.toPath(), ContributionReport.SourceKind.PDF);
        try (PDDocument doc = Loader.loadPDF(pdf)) {
            if (doc.getNumberOfPages() == 0) {
                report.addParseNote("PDF has no pages.");
                return report;
            }
            int pageIndex = 0;
            for (var page : doc.getPages()) {
                pageIndex++;
                PositionCollector collector = new PositionCollector();
                PDFTextStripper stripper = new PDFTextStripper() {
                    @Override
                    protected void writeString(String text, List<TextPosition> positions) {
                        collector.accept(positions);
                    }
                };
                stripper.setStartPage(pageIndex);
                stripper.setEndPage(pageIndex);
                stripper.setSortByPosition(true);
                stripper.getText(doc); // side effect feeds the collector

                List<List<TableExtractor.Word>> lines = collector.toLines();
                TableExtractor.Table table = new TableExtractor().extractFromLines(lines);
                if (!table.found()) {
                    report.addParseNote("PDF page " + pageIndex + ": no schedule header detected, page skipped.");
                    continue;
                }
                ScheduleSheet sheet = toSheet("PAGE " + pageIndex, table);
                if (sheet != null) report.getMonthlySheets().add(sheet);
            }
        }
        return report;
    }

    /** Maps reconstructed table columns onto the canonical entry. */
    static ScheduleSheet toSheet(String sheetName, TableExtractor.Table table) {
        int ss = -1, surname = -1, first = -1, other = -1, salary = -1, contrib = -1, member = -1;
        for (int c = 0; c < table.headers.size(); c++) {
            String h = table.headers.get(c).replace(":", "");
            if (h.contains("MEMBER")) member = setIf(member, c);
            else if (h.contains("SURNAME")) surname = setIf(surname, c);
            else if (h.contains("FIRST")) first = setIf(first, c);
            else if (h.contains("OTHER")) other = setIf(other, c);
            else if (h.contains("SALARY")) salary = setIf(salary, c);
            else if (h.contains("CONTRIBUTION")) contrib = setIf(contrib, c);
            else if (h.startsWith("SS") || h.startsWith("SSNIT")) ss = setIf(ss, c);
        }
        if (ss < 0 && surname < 0) return null;

        List<ContributionEntry> entries = new ArrayList<>();
        int rowNo = 0;
        Double statedSalary = null;
        Double statedContribution = null;
        for (List<String> row : table.rows) {
            rowNo++;
            String ssNo = cell(row, ss);
            String sur = cell(row, surname);
            String fn = cell(row, first);
            String on = cell(row, other);
            Double sal = num(cell(row, salary));
            Double con = num(cell(row, contrib));
            boolean hasIdentity = !ssNo.isBlank() || !sur.isBlank() || !fn.isBlank();
            if (!hasIdentity) {
                if (sal != null && con != null && !entries.isEmpty() && statedContribution == null) {
                    statedSalary = sal;
                    statedContribution = con;
                }
                continue;
            }
            if (sal == null && con == null) continue;
            Integer seq = parseSeq(cell(row, 0));
            entries.add(new ContributionEntry(rowNo, seq, cell(row, member), ssNo, sur, fn, on,
                    "", sal, con, "", ""));
        }
        if (entries.isEmpty()) return null;
        return new ScheduleSheet(sheetName, null, null, sheetName, table.headerLineIndex + 1,
                entries, statedSalary, statedContribution, statedContribution == null ? 0 : table.headerLineIndex + 1 + table.rows.size(),
                new ArrayList<>(), null);
    }

    private static int setIf(int current, int candidate) {
        return current < 0 ? candidate : current;
    }

    private static String cell(List<String> row, int idx) {
        return idx >= 0 && idx < row.size() ? row.get(idx).trim() : "";
    }

    private static Integer parseSeq(String s) {
        try { return Integer.parseInt(s.trim()); } catch (NumberFormatException e) { return null; }
    }

    private static Double num(String s) {
        if (s == null) return null;
        String t = s.replace(",", "").trim();
        if (t.isEmpty()) return null;
        try { return Double.parseDouble(t); } catch (NumberFormatException e) { return null; }
    }

    /** Collects TextPositions into words and words into y-clustered lines. */
    private static final class PositionCollector {
        record Tok(String text, float x0, float x1, float y, float h) {}
        final List<Tok> toks = new ArrayList<>();

        void accept(List<TextPosition> positions) {
            for (TextPosition p : positions) {
                String t = p.getUnicode();
                if (t == null || t.isBlank()) continue;
                float x = p.getXDirAdj();
                float y = p.getYDirAdj();
                float w = p.getWidthDirAdj();
                float h = p.getHeightDir();
                // merge into an open word on the same baseline when gap is small
                Tok last = toks.isEmpty() ? null : toks.get(toks.size() - 1);
                if (last != null && Math.abs(last.y - y) < LINE_TOLERANCE
                        && x - last.x1 < p.getFontSize() * 0.6f) {
                    toks.set(toks.size() - 1, new Tok(last.text + t, last.x0, x + w, y, Math.max(last.h, h)));
                } else {
                    toks.add(new Tok(t.trim(), x, x + w, y, h));
                }
            }
        }

        List<List<TableExtractor.Word>> toLines() {
            List<Tok> sorted = new ArrayList<>(toks);
            sorted.sort((a, b) -> {
                int byY = Float.compare(a.y, b.y);
                return byY != 0 ? byY : Float.compare(a.x0, b.x0);
            });
            List<List<TableExtractor.Word>> lines = new ArrayList<>();
            List<TableExtractor.Word> current = new ArrayList<>();
            float currentY = Float.NaN;
            for (Tok t : sorted) {
                if (!Float.isNaN(currentY) && Math.abs(t.y - currentY) > LINE_TOLERANCE * 1.5f) {
                    if (!current.isEmpty()) lines.add(current);
                    current = new ArrayList<>();
                }
                currentY = t.y;
                current.add(new TableExtractor.Word(t.text, t.x0, t.x1));
            }
            if (!current.isEmpty()) lines.add(current);
            return lines;
        }
    }
}
