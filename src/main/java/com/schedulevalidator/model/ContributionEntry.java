package com.schedulevalidator.model;

/**
 * One member's contribution line as captured from any source (xlsx row, PDF table row, OCR row).
 */
public record ContributionEntry(
        int rowNumber,
        Integer seqNo,
        String memberCode,
        String ssNumber,
        String surname,
        String firstName,
        String otherNames,
        String fullName,
        Double basicSalary,
        Double contribution,
        String description,
        String dateOfBirth) {

    public String displayName() {
        StringBuilder sb = new StringBuilder();
        append(sb, surname);
        append(sb, firstName);
        append(sb, otherNames);
        if (sb.isEmpty()) {
            append(sb, fullName);
        }
        String s = sb.toString().trim();
        return s.isEmpty() ? (ssNumber == null ? "(no name)" : ssNumber) : s;
    }

    private static void append(StringBuilder sb, String part) {
        if (part != null && !part.isBlank()) {
            if (!sb.isEmpty()) {
                sb.append(' ');
            }
            sb.append(part.trim());
        }
    }

    /** Lowercase, whitespace-collapsed name used for matching across reports. */
    public String normalizedName() {
        return displayName().toLowerCase().replaceAll("\\s+", " ").trim();
    }

    /** SS numbers are the primary join key; uppercase, no spaces. */
    public String normalizedSs() {
        return ssNumber == null ? "" : ssNumber.replaceAll("\\s+", "").toUpperCase();
    }
}
