package com.schedulevalidator;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/** Command-line entry point: java ... com.schedulevalidator.HeadlessRunner [options] */
public final class HeadlessRunner {

    public static void main(String[] args) {
        Path reference = null;
        Path upload = null;
        Path outDir = Pipeline.defaultOutDir();
        try {
            for (int i = 0; i < args.length; i++) {
                switch (args[i]) {
                    case "--reference", "-r" -> reference = Path.of(args[++i]);
                    case "--uploaded", "-u" -> upload = Path.of(args[++i]);
                    case "--out", "-o" -> outDir = Path.of(args[++i]);
                    case "--help", "-h" -> {
                        printHelp();
                        return;
                    }
                    default -> throw new IllegalArgumentException("Unknown option: " + args[i]);
                }
            }
            if (reference == null && upload == null) {
                printHelp();
                return;
            }
            List<String> log = new ArrayList<>();
            Pipeline.Result out = Pipeline.run(reference, upload, outDir, Pipeline.resolveTessdataDir(), log);
            log.forEach(System.out::println);
            long errors = 0;
            System.out.println("Done. Report: " + out.reportFile()
                    + (out.convertedFile() != null ? " | Converted: " + out.convertedFile() : ""));
        } catch (Exception e) {
            System.err.println("ERROR: " + e.getMessage());
            System.exit(1);
        }
    }

    private static void printHelp() {
        System.out.println("""
                Schedule Validator - headless mode

                Usage:
                  java -cp <classpath> com.schedulevalidator.HeadlessRunner [options]

                Options:
                  -r, --reference <file>   existing contribution report (.xlsx)
                  -u, --uploaded  <file>   uploaded schedule (.xlsx, .pdf, or image)
                  -o, --out       <dir>    output directory (default: ./out)

                Examples:
                  HeadlessRunner -r Samples/report.xlsx -u uploads/photo.jpg -o out
                  HeadlessRunner -u uploads/march.pdf -o out      (parse+validate+convert only)""");
    }
}
