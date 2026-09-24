package com.schedulevalidator;

import javafx.application.Application;

/**
 * Bare launcher so the jar manifest can point at a class that does not extend
 * javafx.application.Application (avoids the "JavaFX runtime components are missing" error
 * when running from a plain classpath).
 */
public final class Main {
    public static void main(String[] args) {
        if (args.length > 0) {
            HeadlessRunner.main(args);
        } else {
            Application.launch(ValidatorApp.class, args);
        }
    }
}
