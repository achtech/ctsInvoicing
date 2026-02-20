package forecast;


import forecast.by.extcode.ExcelReader;
import forecast.by.extcode.ServiceTeamParser;
import forecast.by.month.service.ExecuteService;
import forecast.by.rate.InputFilesReader;
import forecast.by.rate.InputRowProcessor;
import forecast.by.rate.OutputWriter;
import forecast.by.util.GroupAggregator;
import forecast.by.util.ReferenceData;
import forecast.by.util.ServiceTeam;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class UnifiedMain extends JFrame {

    // Color Palette
    private static final Color PRIMARY_COLOR = new Color(41, 128, 185);
    private static final Color BG_COLOR      = new Color(245, 245, 245);
    private static final Color TEXT_COLOR    = new Color(44, 62, 80);
    private static final Font  MAIN_FONT     = new Font("Segoe UI", Font.PLAIN, 13);
    private static final Font  HEADER_FONT   = new Font("Segoe UI", Font.BOLD, 14);

    public UnifiedMain() {
        setTitle("CTS Invoicing Unified Tool");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1100, 820);
        setMinimumSize(new Dimension(900, 700));
        setLocationRelativeTo(null);
        setBackground(BG_COLOR);
        add(new AllInOnePanel());
    }

    public static void main(String[] args) {
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (Exception e) {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

        UIManager.put("Label.font", MAIN_FONT);
        UIManager.put("Button.font", MAIN_FONT);
        UIManager.put("TextField.font", MAIN_FONT);
        UIManager.put("TextArea.font", new Font("Monospaced", Font.PLAIN, 12));

        SwingUtilities.invokeLater(() -> new UnifiedMain().setVisible(true));
    }

    private static JButton createStyledButton(String text) {
        JButton btn = new JButton(text);
        btn.setBackground(PRIMARY_COLOR);
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.setFont(new Font("Segoe UI", Font.BOLD, 13));
        btn.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(PRIMARY_COLOR.darker(), 1),
                BorderFactory.createEmptyBorder(6, 16, 6, 16)
        ));
        return btn;
    }

    private static TitledBorder createTitledBorder(String title) {
        TitledBorder border = BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(Color.GRAY, 1, true), title);
        border.setTitleFont(HEADER_FONT);
        border.setTitleColor(TEXT_COLOR);
        return border;
    }

    // ==========================================
    // All-in-One Panel
    // ==========================================
    static class AllInOnePanel extends JPanel {

        private final DefaultListModel<File> inputFilesModel  = new DefaultListModel<>();
        private final JList<File>            inputFilesList   = new JList<>(inputFilesModel);
        private final JTextField             targetDirField   = new JTextField();
        private final JSpinner               monthsSpinner    = new JSpinner(new SpinnerNumberModel(3, 1, 12, 1));
        private final JCheckBox              monthsToggle     = new JCheckBox("Enable", true);
        private final JLabel                 inputErrorLabel  = new JLabel(" ");
        private final JLabel                 outputErrorLabel = new JLabel(" ");
        private final JTextArea              logArea          = new JTextArea();

        private final JProgressBar progressBar = new JProgressBar(0, 3);
        private final JLabel       statusLabel = new JLabel("Ready");

        private final JButton runBtn         = createStyledButton("RUN ALL PROCESSES");
        private final JButton openOutputBtn  = createStyledButton("OPEN MAIN OUTPUT FOLDER");

        private File lastMainOutputFolder;

        private static final String HISTORY_PATH = "ctsInvoicing/src/main/resources/history.csv";

        public AllInOnePanel() {
            setLayout(new BorderLayout(10, 10));
            setBorder(new EmptyBorder(15, 15, 15, 15));
            setBackground(BG_COLOR);

            // ── Config panel ─────────────────────────────────────────────────
            JPanel configPanel = new JPanel(new GridBagLayout());
            configPanel.setBackground(Color.WHITE);
            configPanel.setBorder(createTitledBorder("Unified Process Configuration"));

            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(6, 10, 2, 10);
            gbc.fill   = GridBagConstraints.HORIZONTAL;

            // Row 0: label
            gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0; gbc.gridwidth = 1;
            configPanel.add(new JLabel("Input Excel Files:"), gbc);

            // Row 1: file list
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 2;
            gbc.fill    = GridBagConstraints.BOTH;
            gbc.weighty = 1.0;
            inputFilesList.setFixedCellHeight(22);
            JScrollPane listScroll = new JScrollPane(inputFilesList);
            listScroll.setPreferredSize(new Dimension(0, 90));
            listScroll.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
            configPanel.add(listScroll, gbc);
            gbc.weighty = 0;

            // Row 2: input error
            gbc.fill   = GridBagConstraints.HORIZONTAL;
            gbc.gridx  = 0; gbc.gridy = 2; gbc.gridwidth = 2;
            gbc.insets = new Insets(0, 10, 0, 10);
            inputErrorLabel.setForeground(new Color(192, 57, 43));
            inputErrorLabel.setFont(MAIN_FONT.deriveFont(Font.BOLD, 11f));
            configPanel.add(inputErrorLabel, gbc);
            gbc.insets = new Insets(4, 10, 4, 10);

            // Row 3: file action buttons
            gbc.fill   = GridBagConstraints.NONE;
            gbc.gridx  = 0; gbc.gridy = 3; gbc.gridwidth = 2;
            gbc.anchor = GridBagConstraints.EAST;
            JPanel fileButtonsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
            fileButtonsPanel.setOpaque(false);
            JButton addBtn    = createStyledButton("Add Files");
            JButton removeBtn = createStyledButton("Remove Selected");
            JButton clearBtn  = createStyledButton("Clear All");
            addBtn.addActionListener(e    -> selectInputs());
            removeBtn.addActionListener(e -> removeSelectedInputs());
            clearBtn.addActionListener(e  -> clearInputs());
            fileButtonsPanel.add(addBtn);
            fileButtonsPanel.add(removeBtn);
            fileButtonsPanel.add(clearBtn);
            configPanel.add(fileButtonsPanel, gbc);
            gbc.anchor = GridBagConstraints.WEST;

            // Row 4: output dir
            gbc.fill   = GridBagConstraints.HORIZONTAL;
            gbc.gridx  = 0; gbc.gridy = 4; gbc.gridwidth = 1; gbc.weightx = 0;
            configPanel.add(new JLabel("Target Output Directory:"), gbc);

            gbc.gridx = 1; gbc.gridy = 4; gbc.weightx = 1.0;
            JPanel dirPanel = new JPanel(new BorderLayout(8, 0));
            dirPanel.setOpaque(false);
            targetDirField.setEditable(false);
            dirPanel.add(targetDirField, BorderLayout.CENTER);
            JButton selectTargetBtn = createStyledButton("Browse...");
            selectTargetBtn.addActionListener(e -> selectTarget());
            dirPanel.add(selectTargetBtn, BorderLayout.EAST);
            configPanel.add(dirPanel, gbc);

            // Row 5: output error
            gbc.gridx  = 0; gbc.gridy = 5; gbc.gridwidth = 2;
            gbc.insets = new Insets(0, 10, 0, 10);
            outputErrorLabel.setForeground(new Color(192, 57, 43));
            outputErrorLabel.setFont(MAIN_FONT.deriveFont(Font.BOLD, 11f));
            configPanel.add(outputErrorLabel, gbc);
            gbc.insets = new Insets(4, 10, 4, 10);

            // Row 6: months spinner
            gbc.gridx = 0; gbc.gridy = 6; gbc.gridwidth = 1; gbc.weightx = 0;
            configPanel.add(new JLabel("Forecast Months (Month Module):"), gbc);

            gbc.gridx = 1; gbc.gridy = 6; gbc.weightx = 1.0;
            JComponent editor = monthsSpinner.getEditor();
            if (editor instanceof JSpinner.DefaultEditor)
                ((JSpinner.DefaultEditor) editor).getTextField().setColumns(4);
            JPanel spinnerPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
            spinnerPanel.setOpaque(false);
            monthsToggle.setOpaque(false);
            monthsToggle.setFont(MAIN_FONT);
            monthsToggle.addActionListener(e -> monthsSpinner.setEnabled(monthsToggle.isSelected()));
            spinnerPanel.add(monthsToggle);
            spinnerPanel.add(Box.createHorizontalStrut(10));
            spinnerPanel.add(monthsSpinner);
            configPanel.add(spinnerPanel, gbc);

            // Row 7: run button
            gbc.gridx  = 0; gbc.gridy = 7; gbc.gridwidth = 2;
            gbc.fill   = GridBagConstraints.NONE;
            gbc.anchor = GridBagConstraints.CENTER;
            gbc.insets = new Insets(10, 10, 12, 10);
            runBtn.setBackground(new Color(231, 76, 60));
            runBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
            runBtn.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(new Color(192, 57, 43), 1),
                    BorderFactory.createEmptyBorder(8, 28, 8, 28)
            ));
            runBtn.addActionListener(e -> runAll());
            openOutputBtn.setEnabled(false);
            openOutputBtn.addActionListener(e -> openOutputFolder());
            JPanel actionButtonsPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 0));
            actionButtonsPanel.setOpaque(false);
            actionButtonsPanel.add(runBtn);
            actionButtonsPanel.add(openOutputBtn);
            configPanel.add(actionButtonsPanel, gbc);

            // ── Progress panel ───────────────────────────────────────────────
            progressBar.setStringPainted(true);
            progressBar.setString("Idle");
            progressBar.setValue(0);
            progressBar.setPreferredSize(new Dimension(0, 30));
            progressBar.setForeground(new Color(39, 174, 96));
            progressBar.setBackground(new Color(210, 210, 210));
            progressBar.setFont(new Font("Segoe UI", Font.BOLD, 12));

            statusLabel.setFont(MAIN_FONT);
            statusLabel.setForeground(TEXT_COLOR);
            statusLabel.setBorder(new EmptyBorder(2, 4, 4, 4));

            JPanel progressPanel = new JPanel(new BorderLayout(4, 2));
            progressPanel.setBackground(BG_COLOR);
            progressPanel.setBorder(createTitledBorder("Execution Progress"));
            progressPanel.add(progressBar,  BorderLayout.CENTER);
            progressPanel.add(statusLabel,  BorderLayout.SOUTH);
            progressPanel.setPreferredSize(new Dimension(0, 85));

            // ── Log panel ────────────────────────────────────────────────────
            logArea.setEditable(false);
            logArea.setLineWrap(false);
            JScrollPane logScroll = new JScrollPane(logArea);
            logScroll.setBorder(createTitledBorder("Execution Logs"));

            // ── Bottom: progress + logs stacked ─────────────────────────────
            JPanel bottomPanel = new JPanel(new BorderLayout(0, 8));
            bottomPanel.setOpaque(false);
            bottomPanel.add(progressPanel, BorderLayout.NORTH);
            bottomPanel.add(logScroll,     BorderLayout.CENTER);

            // ── Split: config (top) / progress+logs (bottom) ─────────────────
            JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT, configPanel, bottomPanel);
            split.setResizeWeight(0.50);  // config gets ~50%
            split.setDividerSize(7);
            split.setBorder(null);
            split.setOpaque(false);

            add(split, BorderLayout.CENTER);

            loadHistory();
        }

        // ── File helpers ─────────────────────────────────────────────────────

        private void selectInputs() {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION)
                for (File f : chooser.getSelectedFiles())
                    if (!inputFilesModel.contains(f)) inputFilesModel.addElement(f);
        }

        private void removeSelectedInputs() {
            List<File> selected = inputFilesList.getSelectedValuesList();
            if (selected == null || selected.isEmpty()) return;
            for (File f : selected) inputFilesModel.removeElement(f);
        }

        private void clearInputs() {
            inputFilesModel.clear();
        }

        private void selectTarget() {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                File sel = chooser.getSelectedFile();
                targetDirField.setText(sel.getAbsolutePath());
                appendHistory(sel.getAbsolutePath(), (Integer) monthsSpinner.getValue());
            }
        }

        // ── History ──────────────────────────────────────────────────────────

        private void appendHistory(String path, int months) {
            File file = new File(HISTORY_PATH);
            File parent = file.getParentFile();
            if (parent != null && !parent.exists()) parent.mkdirs();
            try (BufferedWriter bw = new BufferedWriter(new FileWriter(file, false))) {
                bw.write(java.time.LocalDateTime.now() + ";" + path + ";" + months);
                bw.newLine();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        private void loadHistory() {
            File file = new File(HISTORY_PATH);
            if (!file.exists()) return;
            String lastLine = null;
            try (BufferedReader br = new BufferedReader(new FileReader(file))) {
                String line;
                while ((line = br.readLine()) != null)
                    if (!line.trim().isEmpty()) lastLine = line;
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (lastLine == null) return;
            String[] parts = lastLine.split(";", -1);
            if (parts.length >= 3) {
                if (!parts[1].isEmpty()) targetDirField.setText(parts[1]);
                try { monthsSpinner.setValue(Integer.parseInt(parts[2])); } catch (NumberFormatException ignored) {}
            }
        }

        // ── Progress / log ───────────────────────────────────────────────────

        /** Update the progress bar. value must be 0..3. */
        private void setProgress(int value, String barLabel, String detail) {
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(value);
                int pct = (value * 100) / progressBar.getMaximum();
                progressBar.setString(pct + "%  —  " + barLabel);
                statusLabel.setText("  ▸  " + detail);
            });
        }

        private void resetProgress() {
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(0);
                progressBar.setString("Starting...");
                statusLabel.setText("  ▸  Initializing");
            });
        }

        private void log(String msg) {
            SwingUtilities.invokeLater(() -> {
                logArea.append(msg + "\n");
                logArea.setCaretPosition(logArea.getDocument().getLength());
            });
        }

        // ── Main execution ───────────────────────────────────────────────────

        private void runAll() {
            inputErrorLabel.setText(" ");
            outputErrorLabel.setText(" ");

            boolean hasError = false;
            if (inputFilesModel.isEmpty()) {
                inputErrorLabel.setText("Please add at least one input Excel file.");
                hasError = true;
            }
            if (targetDirField.getText().isEmpty()) {
                outputErrorLabel.setText("Please select an output directory.");
                hasError = true;
            }
            if (hasError) return;

            File targetDir = new File(targetDirField.getText());
            if (!targetDir.exists() || !targetDir.isDirectory()) {
                JOptionPane.showMessageDialog(this, "Invalid target directory.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            java.time.LocalDateTime now = java.time.LocalDateTime.now();
            String currentMonthStr = now.format(java.time.format.DateTimeFormatter.ofPattern("MMM_yyyy"));
            String runStamp        = now.format(java.time.format.DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

            File mainOutputFolder = new File(targetDir, "forecast_italy_" + currentMonthStr + "_" + runStamp);
            mainOutputFolder.mkdirs();
            lastMainOutputFolder = mainOutputFolder;

            File rateFolder  = new File(mainOutputFolder, "forecast_it_rate_" + currentMonthStr);
            File extFolder   = new File(mainOutputFolder, "forecast_EXT_"     + currentMonthStr);
            File monthFolder = new File(mainOutputFolder, "forecast_month_"   + currentMonthStr);
            rateFolder.mkdirs(); extFolder.mkdirs(); monthFolder.mkdirs();

            int months       = (Integer) monthsSpinner.getValue();
            boolean useManual = monthsToggle.isSelected();
            appendHistory(targetDir.getAbsolutePath(), months);

            List<File> inputs = new ArrayList<>();
            for (int i = 0; i < inputFilesModel.size(); i++) inputs.add(inputFilesModel.get(i));

            runBtn.setEnabled(false);
            runBtn.setText("Running...");
            openOutputBtn.setEnabled(false);
            resetProgress();

            new Thread(() -> {
                log("=== STARTING UNIFIED EXECUTION ===");
                log("Output Folder : " + mainOutputFolder.getAbsolutePath());
                log("Months setting: " + months);

                // ── 1. RATE ──────────────────────────────────────────────────
                setProgress(0, "Step 1/3 — Rate", "Running Forecast By Rate...");
                log("\n[1/3] Running Forecast By Rate...");
                try {
                    ReferenceData referenceData = new ReferenceData();
                    String dataPath = "C:\\Users\\Sanae\\Desktop\\Task_java_excel\\ctsInvoicing\\src\\main\\resources\\Data.xlsx";
                    if (!new File(dataPath).exists()) dataPath = "src/main/resources/Data.xlsx";

                    referenceData.load(dataPath);
                    GroupAggregator   aggregator   = new GroupAggregator();
                    InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
                    InputFilesReader  filesReader  = new InputFilesReader(rowProcessor, aggregator);

                    for (File f : inputs) {
                        try { filesReader.processFile(f.getAbsolutePath()); }
                        catch (Exception e) { log("  - Rate Warning: Failed to process " + f.getName()); }
                    }

                    if (!aggregator.getAggregates().isEmpty()) {
                        OutputWriter writer      = new OutputWriter(referenceData, aggregator);
                        String fullMonth         = now.format(java.time.format.DateTimeFormatter.ofPattern("MMMM"));
                        String rateOut           = new File(rateFolder, "Rate Forecast " + fullMonth + ".xlsx").getAbsolutePath();
                        writer.write(rateOut);
                        log("  > Rate Report created: " + rateOut);
                    } else {
                        log("  - Rate Warning: No valid data found for Rate module.");
                    }
                } catch (Exception e) {
                    log("  ! Rate Module Failed: " + e.getMessage());
                    e.printStackTrace();
                }

                // ── 2. EXT CODE ──────────────────────────────────────────────
                setProgress(1, "Step 2/3 — ExtCode", "Running Forecast By ExtCode...");
                log("\n[2/3] Running Forecast By ExtCode...");
                try {
                    ExcelReader       reader  = new ExcelReader();
                    ServiceTeamParser parser  = new ServiceTeamParser();
                    forecast.by.extcode.ExcelWriter writer = new forecast.by.extcode.ExcelWriter();

                    List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
                    for (File f : inputs) {
                        try { rawItems.addAll(reader.extractRawServiceTeams(f)); }
                        catch (Exception e) { log("  - ExtCode Warning: Failed to process " + f.getName()); }
                    }

                    if (!rawItems.isEmpty()) {
                        List<String> labels = new ArrayList<>();
                        for (ExcelReader.ServiceTeamRaw raw : rawItems) labels.add(raw.getLabel());
                        List<ServiceTeam> parsed = parser.parse(labels);
                        for (int i = 0; i < parsed.size(); i++) {
                            parsed.get(i).setCost(rawItems.get(i).getCost() == null ? "" : String.valueOf(rawItems.get(i).getCost()));
                            parsed.get(i).setStyle(rawItems.get(i).getCost() == null ? null : rawItems.get(i).getStyle());
                        }
                        writer.write(parsed, extFolder);
                        log("  > ExtCode Report created in: " + extFolder.getAbsolutePath());
                    } else {
                        log("  - ExtCode Warning: No valid data found.");
                    }
                } catch (Exception e) {
                    log("  ! ExtCode Module Failed: " + e.getMessage());
                    e.printStackTrace();
                }

                // ── 3. MONTH ─────────────────────────────────────────────────
                setProgress(2, "Step 3/3 — Month", "Running Forecast By Month...");
                log("\n[3/3] Running Forecast By Month...");
                try {
                    for (File f : inputs) {
                        try {
                            int currentMonths = months;
                            if (!useManual) {
                                int detected = countMonthSheets(f);
                                if (detected > 0) {
                                    currentMonths = detected;
                                    log("  - Auto-detected months for " + f.getName() + ": " + currentMonths);
                                } else {
                                    log("  - Warning: No Facturacion sheets found in " + f.getName() + ". Using default: " + months);
                                }
                            }
                            log("  - Processing " + f.getName() + " with " + currentMonths + " months...");
                            ExecuteService.executeScript(f.getAbsolutePath(), monthFolder.getAbsolutePath(), currentMonths);
                        } catch (Exception e) {
                            log("  - Month Warning: Failed to process " + f.getName() + ": " + e.getMessage());
                        }
                    }
                    log("  > Month processing finished.");
                } catch (Exception e) {
                    log("  ! Month Module Critical Error: " + e.getMessage());
                }

                // ── Done ─────────────────────────────────────────────────────
                setProgress(3, "Completed", "All modules finished successfully.");
                log("\n=== EXECUTION COMPLETED ===");

                SwingUtilities.invokeLater(() -> {
                    runBtn.setEnabled(true);
                    runBtn.setText("RUN ALL PROCESSES");
                    if (lastMainOutputFolder != null) {
                        openOutputBtn.setEnabled(true);
                        openOutputBtn.setText("OPEN: " + lastMainOutputFolder.getName());
                    }
                    JOptionPane.showMessageDialog(this,
                            "All processes finished. Check logs for details.\nOutput: "
                            + mainOutputFolder.getAbsolutePath());
                });
            }).start();
        }

        private void openOutputFolder() {
            if (lastMainOutputFolder == null) return;
            if (!lastMainOutputFolder.exists() || !lastMainOutputFolder.isDirectory()) {
                JOptionPane.showMessageDialog(this,
                        "The output folder does not exist:\n" + lastMainOutputFolder.getAbsolutePath(),
                        "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            try {
                if (Desktop.isDesktopSupported()) {
                    Desktop.getDesktop().open(lastMainOutputFolder);
                } else {
                    JOptionPane.showMessageDialog(this,
                            "Opening folders is not supported on this system.",
                            "Error", JOptionPane.ERROR_MESSAGE);
                }
            } catch (IOException e) {
                JOptionPane.showMessageDialog(this,
                        "Failed to open the output folder:\n" + lastMainOutputFolder.getAbsolutePath(),
                        "Error", JOptionPane.ERROR_MESSAGE);
            }
        }

        // ── Sheet counter ─────────────────────────────────────────────────────

        private int countMonthSheets(File f) {
            int count = 0;
            try (org.apache.poi.ss.usermodel.Workbook wb = org.apache.poi.ss.usermodel.WorkbookFactory.create(f)) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    String n = wb.getSheetName(i).toLowerCase()
                            .replace("á","a").replace("é","e")
                            .replace("í","i").replace("ó","o").replace("ú","u");
                    if (n.contains("facturacion")) count++;
                }
            } catch (Exception e) {
                log("  - Error counting sheets in " + f.getName() + ": " + e.getMessage());
            }
            return count;
        }
    }
}
