package com.schedulevalidator;

import com.schedulevalidator.model.ValidationIssue;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.nio.file.Files;
import java.nio.file.Path;
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

    private Label refLabel;
    private Label upLabel;
    private Button validateButton;
    private Button exportReportButton;
    private Button exportConvertedButton;
    private ProgressIndicator progress;
    private Label statusLabel;
    private TableView<ValidationIssue> table;

    @Override
    public void start(Stage stage) {
        stage.setTitle("Schedule Validator");

        BorderPane root = new BorderPane();
        root.setTop(buildInputPanel(stage));
        root.setCenter(buildResultsTable());
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
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Contribution report (xlsx)", "*.xlsx"));
            Path p = choose(stage, fc);
            if (p != null) {
                referenceFile = p;
                refLabel.setText(p.toString());
            }
        });

        upLabel = new Label("no file selected");
        Button pickUp = new Button("Browse...");
        pickUp.setOnAction(e -> {
            FileChooser fc = new FileChooser();
            fc.setTitle("Choose the uploaded schedule");
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter(
                    "Schedule (xlsx, pdf, image)", "*.xlsx", "*.pdf",
                    "*.png", "*.jpg", "*.jpeg", "*.tif", "*.tiff", "*.bmp"));
            Path p = choose(stage, fc);
            if (p != null) {
                uploadFile = p;
                upLabel.setText(p.toString());
            }
        });

        validateButton = new Button("Validate");
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

        HBox actions = new HBox(10, validateButton, progress);
        actions.setAlignment(Pos.CENTER_LEFT);
        grid.add(actions, 2, 2, 2, 1);

        return grid;
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

    private javafx.scene.Parent buildStatusBar() {
        statusLabel = new Label("Ready.");
        exportReportButton = new Button("Export validation report");
        exportReportButton.setDisable(true);
        exportReportButton.setOnAction(e -> copyOut(lastReportFile, "validation_report.xlsx"));
        exportConvertedButton = new Button("Save converted .xlsx");
        exportConvertedButton.setDisable(true);
        exportConvertedButton.setOnAction(e -> copyOut(lastConvertedFile, "converted_schedule.xlsx"));
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox box = new HBox(12, statusLabel, spacer, exportReportButton, exportConvertedButton);
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
        validateButton.setDisable(true);
        progress.setVisible(true);
        statusLabel.setText("Working...");

        Path ref = referenceFile;
        Path up = uploadFile;
        Path outDir = Pipeline.defaultOutDir();
        Path tessdata = Pipeline.resolveTessdataDir();

        Task<Pipeline.Result> task = new Task<>() {
            @Override
            protected Pipeline.Result call() throws Exception {
                List<String> log = new java.util.ArrayList<>();
                Pipeline.Result result = Pipeline.run(ref, up, outDir, tessdata, log);
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
            statusLabel.setText("Done. " + errors + " error(s), " + warnings + " warning(s), "
                    + result.issues().size() + " finding(s) total. Outputs in " + outDir.toAbsolutePath());
            lastReportFile = result.reportFile();
            lastConvertedFile = result.convertedFile();
            exportReportButton.setDisable(false);
            exportConvertedButton.setDisable(lastConvertedFile == null);
        });
        task.setOnFailed(e -> {
            progress.setVisible(false);
            validateButton.setDisable(false);
            Throwable ex = task.getException();
            statusLabel.setText("Failed: " + (ex == null ? "unknown error" : ex.getMessage()));
        });
        Thread worker = new Thread(task, "validation-worker");
        worker.setDaemon(true);
        worker.start();
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
