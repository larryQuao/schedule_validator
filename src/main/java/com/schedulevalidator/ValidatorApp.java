package com.schedulevalidator;

import com.schedulevalidator.convert.SheetMerger;
import com.schedulevalidator.model.ValidationIssue;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * Desktop UI: pick a reference report and/or an uploaded schedule, validate, browse findings,
 * export the validation report and (for pdf/image uploads) the converted .xlsx.
 */
public class ValidatorApp extends javafx.application.Application {

    private Path referenceFile;
    private Path uploadFile;
    private Path lastReportFile;
    private Path lastConvertedFile;
    private Path lastUpdatedFile;
    private Path lastFinalFile;

    private Label refLabel;
    private Label upLabel;
    private ComboBox<String> sheetBox;
    private Button validateButton;
    private Button exportReportButton;
    private Button exportConvertedButton;
    private Button exportUpdatedButton;
    private Button exportFinalButton;
    private ProgressIndicator progress;
    private Label statusLabel;
    private TableView<ValidationIssue> table;
    private TableView<SheetMerger.RecordRow> recordsTable;
    private Tab recordsTab;
    private javafx.scene.control.TextField recordsSearch;
    private javafx.collections.transformation.FilteredList<SheetMerger.RecordRow> filteredRecords;
    private final javafx.collections.ObservableList<SheetMerger.RecordRow> masterRecords =
            javafx.collections.FXCollections.observableArrayList();

    @Override
    public void start(Stage stage) {
        stage.setTitle("Schedule Validator");

        BorderPane root = new BorderPane();
        root.setTop(buildInputPanel(stage));

        Tab findingsTab = new Tab("Findings", buildResultsTable());
        findingsTab.setClosable(false);
        recordsTab = new Tab("Final Records", buildRecordsTable());
        recordsTab.setClosable(false);
        TabPane tabs = new TabPane(findingsTab, recordsTab);
        root.setCenter(tabs);

        root.setBottom(buildStatusBar());

        Scene scene = new Scene(root, 1200, 750);
        var css = ValidatorApp.class.getResource("/css/app.css");
        if (css != null) scene.getStylesheets().add(css.toExternalForm());
        stage.setScene(scene);
        stage.show();
    }

    // ------------------------------------------------------------------ panels

    private javafx.scene.Parent buildInputPanel(Stage stage) {
        GridPane grid = new GridPane();
        grid.setHgap(10);
        grid.setVgap(8);
        grid.setPadding(new Insets(14));
        ColumnConstraints c1 = new ColumnConstraints();
        c1.setHgrow(Priority.ALWAYS);
        grid.getColumnConstraints().addAll(new ColumnConstraints(120), c1,
                new ColumnConstraints(90), new ColumnConstraints(200));

        refLabel = new Label("no file selected");
        Button pickRef = new Button("Browse...");
        pickRef.setOnAction(e -> {
            FileChooser fc = new FileChooser();
            fc.setTitle("Choose the existing contribution report");
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter(
                    "Contribution report (xlsx, xls)", "*.xlsx", "*.xls", "*.xlsm"));
            Path p = choose(stage, fc);
            if (p != null) {
                referenceFile = p;
                refLabel.setText(p.toString());
                loadSheetNames(p);
            }
        });

        upLabel = new Label("no file selected");
        Button pickUp = new Button("Browse...");
        pickUp.setOnAction(e -> {
            FileChooser fc = new FileChooser();
            fc.setTitle("Choose the uploaded schedule");
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter(
                    "Schedule (xlsx, xls, pdf, image)", "*.xlsx", "*.xls", "*.xlsm", "*.pdf",
                    "*.png", "*.jpg", "*.jpeg", "*.tif", "*.tiff", "*.bmp"));
            Path p = choose(stage, fc);
            if (p != null) {
                uploadFile = p;
                upLabel.setText(p.toString());
            }
        });

        sheetBox = new ComboBox<>();
        sheetBox.setDisable(true);
        sheetBox.setPrefWidth(280);
        sheetBox.setTooltip(new Tooltip("Sheet of the contribution report to compare against"));

        validateButton = new Button("Validate & merge new month");
        validateButton.setStyle("-fx-font-weight: bold;");
        validateButton.setOnAction(e -> runValidation());

        progress = new ProgressIndicator();
        progress.setVisible(false);
        progress.setPrefSize(24, 24);

        grid.add(new Label("Reference report"), 0, 0);
        grid.add(refLabel, 1, 0);
        grid.add(pickRef, 2, 0);
        grid.add(new Label("Uploaded schedule"), 0, 1);
        grid.add(upLabel, 1, 1);
        grid.add(pickUp, 2, 1);
        grid.add(new Label("Reference sheet"), 0, 2);
        grid.add(sheetBox, 1, 2);

        HBox actions = new HBox(10, validateButton, progress);
        actions.setAlignment(Pos.CENTER_LEFT);
        grid.add(actions, 2, 2, 2, 1);

        return grid;
    }

    /** Lists the workbook's sheets and pre-selects the latest monthly one. */
    private void loadSheetNames(Path workbookFile) {
        try (Workbook wb = WorkbookFactory.create(
                new java.io.ByteArrayInputStream(java.nio.file.Files.readAllBytes(workbookFile)))) {
            List<String> names = new ArrayList<>();
            wb.sheetIterator().forEachRemaining(s -> names.add(s.getSheetName()));
            sheetBox.setItems(FXCollections.observableArrayList(names));
            sheetBox.setDisable(names.isEmpty());
            String monthly = null;
            for (String n : names) {
                if (n.toUpperCase().matches(".*(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*.*(19|20)\\d\\d.*")) {
                    monthly = n;
                }
            }
            sheetBox.getSelectionModel().select(monthly != null ? monthly : names.get(names.size() - 1));
            statusLabel.setText("Loaded " + names.size() + " sheets - pick the sheet to compare against.");
        } catch (Exception ex) {
            sheetBox.getItems().clear();
            sheetBox.setDisable(true);
            statusLabel.setText("Could not read sheets: " + ex.getMessage());
        }
    }

    @SuppressWarnings("unchecked")
    private TableView<ValidationIssue> buildResultsTable() {
        table = new TableView<>();

        TableColumn<ValidationIssue, String> sev = new TableColumn<>("Severity");
        sev.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().severity().toString()));
        sev.setPrefWidth(90);
        sev.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty || item == null ? null : item);
                if (!empty && item != null) {
                    getStyleClass().removeAll("sev-error", "sev-warning", "sev-info");
                    getStyleClass().add(switch (item) {
                        case "ERROR" -> "sev-error";
                        case "WARNING" -> "sev-warning";
                        default -> "sev-info";
                    });
                }
            }
        });

        TableColumn<ValidationIssue, String> source = new TableColumn<>("Source");
        source.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().source()));
        source.setPrefWidth(100);

        TableColumn<ValidationIssue, String> sheet = new TableColumn<>("Sheet");
        sheet.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().sheet()));
        sheet.setPrefWidth(150);

        TableColumn<ValidationIssue, String> loc = new TableColumn<>("Location");
        loc.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().location()));
        loc.setPrefWidth(130);

        TableColumn<ValidationIssue, String> rule = new TableColumn<>("Rule");
        rule.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().rule()));
        rule.setPrefWidth(180);

        TableColumn<ValidationIssue, String> msg = new TableColumn<>("Message");
        msg.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().message()));
        msg.setPrefWidth(540);

        table.getColumns().addAll(sev, source, sheet, loc, rule, msg);
        table.setPlaceholder(new Label("Pick files and press Validate."));
        table.setRowFactory(tv -> {
            TableRow<ValidationIssue> row = new TableRow<>();
            row.itemProperty().addListener((obs, old, item) ->
                    row.setTooltip(item == null ? null : new Tooltip(item.message())));
            return row;
        });
        return table;
    }

    /** The merged schedule rows: CARRIED / NEW (green) / REMOVED (red), with live search. */
    @SuppressWarnings("unchecked")
    private javafx.scene.Parent buildRecordsTable() {
        recordsTable = new TableView<>();
        filteredRecords = new javafx.collections.transformation.FilteredList<>(masterRecords, r -> true);
        var sorted = new javafx.collections.transformation.SortedList<>(filteredRecords);
        sorted.comparatorProperty().bind(recordsTable.comparatorProperty());
        recordsTable.setItems(sorted);

        TableColumn<SheetMerger.RecordRow, String> st = new TableColumn<>("Status");
        st.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().status()));
        st.setPrefWidth(100);
        st.setCellFactory(col -> new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty || item == null ? null : item);
                if (!empty && item != null) {
                    getStyleClass().removeAll("sev-error", "sev-warning", "rec-new", "rec-removed");
                    getStyleClass().add(switch (item) {
                        case "NEW" -> "rec-new";
                        case "REMOVED" -> "rec-removed";
                        default -> "sev-info";
                    });
                }
            }
        });

        TableColumn<SheetMerger.RecordRow, String> sn = new TableColumn<>("S/N");
        sn.setCellValueFactory(c -> new SimpleStringProperty(
                c.getValue().seq() == null ? "" : String.valueOf(c.getValue().seq())));
        sn.setPrefWidth(50);

        TableColumn<SheetMerger.RecordRow, String> mc = new TableColumn<>("Member Code");
        mc.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().memberCode()));
        mc.setPrefWidth(140);

        TableColumn<SheetMerger.RecordRow, String> ss = new TableColumn<>("SS No");
        ss.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().ssNumber()));
        ss.setPrefWidth(140);

        TableColumn<SheetMerger.RecordRow, String> surname = new TableColumn<>("Surname");
        surname.setCellValueFactory(c -> new SimpleStringProperty(nz(c.getValue().surname())));
        surname.setPrefWidth(130);

        TableColumn<SheetMerger.RecordRow, String> first = new TableColumn<>("Firstname");
        first.setCellValueFactory(c -> new SimpleStringProperty(nz(c.getValue().firstName())));
        first.setPrefWidth(120);

        TableColumn<SheetMerger.RecordRow, String> other = new TableColumn<>("Other Names");
        other.setCellValueFactory(c -> new SimpleStringProperty(nz(c.getValue().otherNames())));
        other.setPrefWidth(130);

        TableColumn<SheetMerger.RecordRow, String> sal = new TableColumn<>("Basic Salary");
        sal.setCellValueFactory(c -> new SimpleStringProperty(
                c.getValue().basicSalary() == null ? "" : String.format("%,.2f", c.getValue().basicSalary())));
        sal.setPrefWidth(110);

        TableColumn<SheetMerger.RecordRow, String> con = new TableColumn<>("5% Contribution");
        con.setCellValueFactory(c -> new SimpleStringProperty(
                c.getValue().contribution() == null ? "" : String.format("%,.2f", c.getValue().contribution())));
        con.setPrefWidth(120);

        recordsTable.getColumns().addAll(st, sn, mc, ss, surname, first, other, sal, con);
        recordsTable.setPlaceholder(new Label("Run a validation & merge to see the final records here."));

        recordsSearch = new TextField();
        recordsSearch.setPromptText("Search by name or SS number...");
        recordsSearch.textProperty().addListener((obs, ov, nv) -> applyRecordsFilter());
        HBox bar = new HBox(8, new Label("Search:"), recordsSearch);
        bar.setAlignment(Pos.CENTER_LEFT);
        bar.setPadding(new Insets(6));
        HBox.setHgrow(recordsSearch, Priority.ALWAYS);
        VBox box = new VBox(bar, recordsTable);
        VBox.setVgrow(recordsTable, Priority.ALWAYS);
        return box;
    }

    /** Live filter over the final records: any name part, full name, or SS number. */
    private void applyRecordsFilter() {
        String q = recordsSearch.getText() == null ? "" : recordsSearch.getText().trim().toLowerCase();
        filteredRecords.setPredicate(r -> q.isEmpty()
                || contains(r.surname(), q) || contains(r.firstName(), q)
                || contains(r.otherNames(), q) || contains(r.fullName(), q)
                || contains(r.ssNumber(), q));
    }

    private static boolean contains(String value, String query) {
        return value != null && value.toLowerCase().contains(query);
    }

    private static String nz(String s) {
        return s == null ? "" : s;
    }

    private javafx.scene.Parent buildStatusBar() {
        statusLabel = new Label("Ready.");
        exportReportButton = new Button("Export validation report");
        exportReportButton.setDisable(true);
        exportReportButton.setOnAction(e -> copyOut(lastReportFile, "validation_report.xlsx"));
        exportConvertedButton = new Button("Save converted .xlsx");
        exportConvertedButton.setDisable(true);
        exportConvertedButton.setOnAction(e -> copyOut(lastConvertedFile, "converted_schedule.xlsx"));
        exportUpdatedButton = new Button("Save updated report");
        exportUpdatedButton.setDisable(true);
        exportUpdatedButton.setOnAction(e -> copyOut(lastUpdatedFile, "updated_report.xlsx"));
        exportFinalButton = new Button("Download final report");
        exportFinalButton.setDisable(true);
        exportFinalButton.setStyle("-fx-font-weight: bold;");
        exportFinalButton.setOnAction(e -> copyOut(lastFinalFile, "final_validation_report.xlsx"));
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox box = new HBox(12, statusLabel, spacer, exportReportButton, exportConvertedButton,
                exportUpdatedButton, exportFinalButton);
        box.setAlignment(Pos.CENTER_LEFT);
        box.setPadding(new Insets(10));
        return box;
    }

    // ------------------------------------------------------------------ actions

    private Path choose(Stage stage, FileChooser fc) {
        var f = fc.showOpenDialog(stage);
        return f == null ? null : f.toPath();
    }

    private void runValidation() {
        if (referenceFile == null && uploadFile == null) {
            statusLabel.setText("Pick at least one file first.");
            return;
        }
        boolean mergeMode = referenceFile != null && uploadFile != null && !sheetBox.isDisabled();
        if (mergeMode && sheetBox.getSelectionModel().isEmpty()) {
            statusLabel.setText("Select the reference sheet to compare against.");
            return;
        }
        validateButton.setDisable(true);
        progress.setVisible(true);
        statusLabel.setText("Working...");

        Path ref = referenceFile;
        Path up = uploadFile;
        String sheet = mergeMode ? sheetBox.getValue() : null;
        Path outDir = Pipeline.defaultOutDir();
        Path tessdata = Pipeline.resolveTessdataDir();

        Task<Pipeline.Result> task = new Task<>() {
            @Override
            protected Pipeline.Result call() throws Exception {
                List<String> log = new ArrayList<>();
                Pipeline.Result result = (sheet != null)
                        ? Pipeline.runSheetMerge(ref, sheet, up, outDir, tessdata, log)
                        : Pipeline.run(ref, up, outDir, tessdata, log);
                for (String line : log) {
                    updateMessage(line);
                }
                return result;
            }
        };
        task.messageProperty().addListener((obs, old, line) ->
                Platform.runLater(() -> statusLabel.setText(line)));
        task.setOnSucceeded(e -> {
            Pipeline.Result result = task.getValue();
            progress.setVisible(false);
            validateButton.setDisable(false);
            table.setItems(FXCollections.observableArrayList(result.issues()));
            long errors = result.issues().stream().filter(i -> i.severity() == ValidationIssue.Severity.ERROR).count();
            long warnings = result.issues().stream().filter(i -> i.severity() == ValidationIssue.Severity.WARNING).count();
            statusLabel.setText("Done. " + errors + " error(s), " + warnings + " warning(s). Outputs in "
                    + outDir.toAbsolutePath());
            lastReportFile = result.reportFile();
            lastConvertedFile = result.convertedFile();
            lastUpdatedFile = result.updatedWorkbook();
            lastFinalFile = result.finalReportFile();
            exportReportButton.setDisable(false);
            exportConvertedButton.setDisable(lastConvertedFile == null);
            exportUpdatedButton.setDisable(lastUpdatedFile == null);
            exportFinalButton.setDisable(lastFinalFile == null);
            if (result.records() != null) {
                masterRecords.setAll(result.records());
                Platform.runLater(() -> recordsTab.getTabPane().getSelectionModel().select(recordsTab));
            }
        });
        task.setOnFailed(e -> {
            progress.setVisible(false);
            validateButton.setDisable(false);
            Throwable ex = task.getException();
            String msg = ex == null ? "Unknown error" : String.valueOf(ex.getMessage());
            statusLabel.setText("Failed: " + msg);
            showErrorDialog(msg, ex);
        });
        Thread worker = new Thread(task, "validation-worker");
        worker.setDaemon(true);
        worker.start();
    }

    /** Modal error dialogue; adds a close-Excel hint for file-lock failures. */
    private void showErrorDialog(String message, Throwable ex) {
        boolean lock = message != null && (message.contains("being used by another process")
                || message.contains("locked on the network") || message.contains("It is already in use"));
        Alert alert = new Alert(Alert.AlertType.ERROR,
                message + (lock ? "\n\nThe file may be open in Excel — close it there and try again,"
                        + " or save a local copy and pick that." : ""),
                ButtonType.OK);
        alert.setHeaderText(lock ? "File is in use elsewhere" : "Validation failed");
        alert.getDialogPane().setMinWidth(600);
        if (ex != null) {
            java.io.StringWriter sw = new java.io.StringWriter();
            ex.printStackTrace(new java.io.PrintWriter(sw));
            javafx.scene.control.TextArea details = new javafx.scene.control.TextArea(sw.toString());
            details.setEditable(false);
            details.setWrapText(true);
            details.setMaxHeight(220);
            alert.getDialogPane().setExpandableContent(details);
        }
        alert.showAndWait();
    }

    private void copyOut(Path source, String suggestedName) {
        if (source == null || !Files.isRegularFile(source)) return;
        FileChooser fc = new FileChooser();
        fc.setTitle("Save as");
        fc.setInitialFileName(suggestedName);
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel workbook", "*.xlsx"));
        var target = fc.showSaveDialog(table.getScene().getWindow());
        if (target == null) return;
        try {
            Files.copy(source, target.toPath(), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            statusLabel.setText("Saved: " + target);
        } catch (Exception ex) {
            statusLabel.setText("Save failed: " + ex.getMessage());
        }
    }
}
