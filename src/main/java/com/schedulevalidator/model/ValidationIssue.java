package com.schedulevalidator.model;

/** One finding produced by the validation engine or the reference-vs-upload comparison. */
public record ValidationIssue(Severity severity, String source, String sheet, String location,
                              String rule, String message) {

    public enum Severity { ERROR, WARNING, INFO }

    public static ValidationIssue error(String source, String sheet, String location, String rule, String message) {
        return new ValidationIssue(Severity.ERROR, source, sheet, location, rule, message);
    }

    public static ValidationIssue warning(String source, String sheet, String location, String rule, String message) {
        return new ValidationIssue(Severity.WARNING, source, sheet, location, rule, message);
    }

    public static ValidationIssue info(String source, String sheet, String location, String rule, String message) {
        return new ValidationIssue(Severity.INFO, source, sheet, location, rule, message);
    }
}
