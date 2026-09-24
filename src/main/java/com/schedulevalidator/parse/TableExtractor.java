package com.schedulevalidator.parse;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

/**
 * Reconstructs a schedule table from positioned words (from PDF text positions or OCR bounding
 * boxes). Finds the header line by keyword voting ("ss", "surname", "salary", "contribution"),
 * derives column x-ranges from the header words, then assigns each data word to its column by
 * x-proximity. This mirrors how the reference workbook lays out a schedule, so converted output
 * keeps the same column semantics.
 */
public final class TableExtractor {

    public record Word(String text, float x0, float x1) {}

    /** Result of extraction for one table. */
    public static final class Table {
        public int headerLineIndex = -1;
        public final List<String> headers = new ArrayList<>();
        public final List<List<String>> rows = new ArrayList<>();

        public boolean found() { return headerLineIndex >= 0; }
    }

    private static final String[] HEADER_HINTS = {"SS", "SURNAME", "SALARY", "CONTRIBUTION", "MEMBER"};

    /**
     * @param lines one visual line per inner list, each line's words sorted left-to-right
     */
    public Table extractFromLines(List<List<Word>> lines) {
        Table table = new Table();
        int headerIdx = -1;
        int bestScore = 2; // need at least 3 hint hits
        for (int i = 0; i < lines.size(); i++) {
            String joined = String.join(" ", lines.get(i).stream()
                    .map(w -> w.text().toUpperCase(Locale.ROOT)).toList());
            int score = 0;
            for (String h : HEADER_HINTS) if (joined.contains(h)) score++;
            if (score > bestScore) { bestScore = score; headerIdx = i; }
        }
        if (headerIdx < 0) return table;
        table.headerLineIndex = headerIdx;
        List<Word> headerWords = mergeHeaderTokens(new ArrayList<>(lines.get(headerIdx)));
        for (Word w : headerWords) table.headers.add(w.text().toUpperCase(Locale.ROOT));
        for (int i = headerIdx + 1; i < lines.size(); i++) {
            List<String> row = assignToColumns(lines.get(i), headerWords);
            if (row.stream().anyMatch(s -> !s.isBlank())) table.rows.add(row);
        }
        return table;
    }

    /** Merges header words that sit close together into single column tokens ("Member" + "Code"). */
    static List<Word> mergeHeaderTokens(List<Word> headerWords) {
        headerWords.sort(Comparator.comparingDouble(Word::x0));
        List<Word> out = new ArrayList<>();
        float avgWidth = (float) headerWords.stream().mapToDouble(w -> w.x1() - w.x0()).average().orElse(10);
        for (Word w : headerWords) {
            if (!out.isEmpty()) {
                Word prev = out.get(out.size() - 1);
                if (w.x0() - prev.x1() < avgWidth * 0.6f) {
                    out.set(out.size() - 1, new Word(prev.text() + " " + w.text(), prev.x0(), w.x1()));
                    continue;
                }
            }
            out.add(w);
        }
        return out;
    }

    static List<String> assignToColumns(List<Word> words, List<Word> headerColumns) {
        String[] row = new String[headerColumns.size()];
        List<Word> sorted = new ArrayList<>(words);
        sorted.sort(Comparator.comparingDouble(Word::x0));
        for (Word w : sorted) {
            float center = (w.x0() + w.x1()) / 2f;
            int best = 0;
            float bestDist = Float.MAX_VALUE;
            for (int c = 0; c < headerColumns.size(); c++) {
                Word h = headerColumns.get(c);
                float hCenter = (h.x0() + h.x1()) / 2f;
                float halfW = (h.x1() - h.x0()) / 2f;
                float dist = Math.max(0, Math.abs(center - hCenter) - halfW);
                if (dist < bestDist) { bestDist = dist; best = c; }
            }
            row[best] = row[best] == null ? w.text() : row[best] + " " + w.text();
        }
        List<String> out = new ArrayList<>();
        for (String s : row) out.add(s == null ? "" : s);
        return out;
    }
}
