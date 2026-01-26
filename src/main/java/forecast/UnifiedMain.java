package forecast;

import forecast.by.rate.service.InputFilesReader;
import forecast.by.rate.service.InputRowProcessor;
import forecast.by.rate.service.OutputWriter;
import forecast.by.rate.util.GroupAggregator;
import forecast.by.rate.util.ReferenceData;

import forecast.by.month.service.ExecuteService;

import forecast.by.extcode.service.ExcelReader;
import forecast.by.extcode.service.ServiceTeamParser;
import forecast.by.extcode.util.ServiceTeam;

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
    private static final Color PRIMARY_COLOR = new Color(41, 128, 185); // Blue
    private static final Color ACCENT_COLOR = new Color(52, 152, 219); // Lighter Blue
    private static final Color BG_COLOR = new Color(245, 245, 245); // Light Gray
    private static final Color TEXT_COLOR = new Color(44, 62, 80); // Dark Blue/Gray
    private static final Font MAIN_FONT = new Font("Segoe UI", Font.PLAIN, 14);
    private static final Font HEADER_FONT = new Font("Segoe UI", Font.BOLD, 16);

    public UnifiedMain() {
        setTitle("CTS Invoicing Unified Tool");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 700);
        setLocationRelativeTo(null);
        setBackground(BG_COLOR);

        // Custom Tabbed Pane
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setFont(MAIN_FONT);
        
        tabbedPane.addTab("All-in-One Process", new AllInOnePanel());
        tabbedPane.addTab("Forecast By Rate", new RatePanel());
        tabbedPane.addTab("Forecast By Month", new MonthPanel());
        tabbedPane.addTab("Forecast By ExtCode", new ExtCodePanel());

        add(tabbedPane);
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
        
        // Global UI Customization
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
        btn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        btn.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(PRIMARY_COLOR.darker(), 1),
                BorderFactory.createEmptyBorder(8, 20, 8, 20)
        ));
        return btn;
    }

    private static TitledBorder createTitledBorder(String title) {
        TitledBorder border = BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(Color.GRAY, 1, true),
                title
        );
        border.setTitleFont(HEADER_FONT);
        border.setTitleColor(TEXT_COLOR);
        return border;
    }

    // ==========================================
    // 0. All-in-One Panel
    // ==========================================
    static class AllInOnePanel extends JPanel {
        private final DefaultListModel<File> inputFilesModel = new DefaultListModel<>();
        private final JList<File> inputFilesList = new JList<>(inputFilesModel);
        private final JTextField targetDirField = new JTextField();
        private final JSpinner monthsSpinner = new JSpinner(new SpinnerNumberModel(3, 1, 12, 1));
        private final JTextArea logArea = new JTextArea();

        public AllInOnePanel() {
            setLayout(new BorderLayout(15, 15));
            setBorder(new EmptyBorder(20, 20, 20, 20));
            setBackground(BG_COLOR);

            // Configuration Panel
            JPanel configPanel = new JPanel(new GridBagLayout());
            configPanel.setBackground(Color.WHITE);
            configPanel.setBorder(createTitledBorder("Unified Process Configuration"));

            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 10, 10, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            // Row 0: Input Files
            gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0;
            configPanel.add(new JLabel("Input Excel Files:"), gbc);

            gbc.gridx = 1; gbc.gridy = 0; gbc.weightx = 1.0;
            JButton selectInputsBtn = createStyledButton("Select Files");
            selectInputsBtn.addActionListener(e -> selectInputs());
            configPanel.add(selectInputsBtn, gbc);

            // Row 1: List
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 0.5;
            gbc.ipady = 50;
            JScrollPane listScroll = new JScrollPane(inputFilesList);
            listScroll.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
            configPanel.add(listScroll, gbc);
            gbc.ipady = 0;

            // Row 2: Target Directory
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weighty = 0;
            gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 1;
            configPanel.add(new JLabel("Target Output Directory:"), gbc);

            gbc.gridx = 1; gbc.gridy = 2;
            JPanel dirPanel = new JPanel(new BorderLayout(10, 0));
            dirPanel.setOpaque(false);
            targetDirField.setEditable(false);
            dirPanel.add(targetDirField, BorderLayout.CENTER);
            JButton selectTargetBtn = createStyledButton("Browse...");
            selectTargetBtn.addActionListener(e -> selectTarget());
            dirPanel.add(selectTargetBtn, BorderLayout.EAST);
            configPanel.add(dirPanel, gbc);

            // Row 3: Months (for Forecast By Month)
            gbc.gridx = 0; gbc.gridy = 3; gbc.weightx = 0;
            configPanel.add(new JLabel("Forecast Months (for Month Module):"), gbc);

            gbc.gridx = 1; gbc.gridy = 3; gbc.weightx = 1.0;
            JPanel spinnerPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
            spinnerPanel.setOpaque(false);
            spinnerPanel.add(monthsSpinner);
            configPanel.add(spinnerPanel, gbc);

            // Row 4: Execute
            gbc.gridx = 0; gbc.gridy = 4; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.NONE;
            gbc.anchor = GridBagConstraints.CENTER;
            JButton runBtn = createStyledButton("RUN ALL PROCESSES");
            runBtn.setBackground(new Color(231, 76, 60)); // Red/Orange for emphasis
            runBtn.addActionListener(e -> runAll());
            configPanel.add(runBtn, gbc);

            add(configPanel, BorderLayout.NORTH);

            // Logs
            logArea.setEditable(false);
            JScrollPane logScroll = new JScrollPane(logArea);
            logScroll.setBorder(createTitledBorder("Execution Logs"));
            add(logScroll, BorderLayout.CENTER);
        }

        private void selectInputs() {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                for (File f : chooser.getSelectedFiles()) {
                    if (!inputFilesModel.contains(f)) {
                        inputFilesModel.addElement(f);
                    }
                }
            }
        }

        private void selectTarget() {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                targetDirField.setText(chooser.getSelectedFile().getAbsolutePath());
            }
        }

        private void log(String msg) {
            SwingUtilities.invokeLater(() -> {
                logArea.append(msg + "\n");
                logArea.setCaretPosition(logArea.getDocument().getLength());
            });
        }

        private void runAll() {
            if (inputFilesModel.isEmpty() || targetDirField.getText().isEmpty()) {
                JOptionPane.showMessageDialog(this, "Please select input files and output directory.", "Warning", JOptionPane.WARNING_MESSAGE);
                return;
            }

            File targetDir = new File(targetDirField.getText());
            if (!targetDir.exists() || !targetDir.isDirectory()) {
                 JOptionPane.showMessageDialog(this, "Invalid target directory.", "Error", JOptionPane.ERROR_MESSAGE);
                 return;
            }

            int months = (Integer) monthsSpinner.getValue();
            List<File> inputs = new ArrayList<>();
            for(int i=0; i<inputFilesModel.size(); i++) inputs.add(inputFilesModel.get(i));

            new Thread(() -> {
                log("=== STARTING UNIFIED EXECUTION ===");
                
                // 1. RATE MODULE
                try {
                    log("\n[1/3] Running Forecast By Rate...");
                    ReferenceData referenceData = new ReferenceData();
                    String dataPath = "C:\\Users\\Sanae\\Desktop\\Task_java_excel\\ctsInvoicing\\src\\main\\resources\\Data.xlsx";
                    if (!new File(dataPath).exists()) dataPath = "src/main/resources/Data.xlsx";
                    
                    referenceData.load(dataPath);
                    GroupAggregator aggregator = new GroupAggregator();
                    InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
                    InputFilesReader filesReader = new InputFilesReader(rowProcessor, aggregator);

                    for (File f : inputs) {
                        try {
                            filesReader.processFile(f.getAbsolutePath());
                        } catch (Exception e) {
                            log("  - Rate Warning: Failed to process " + f.getName() + " (might be wrong format for Rate)");
                        }
                    }

                    // Print raw data grouped by user
                    rowProcessor.printRawRates();

                    if (!aggregator.getAggregates().isEmpty()) {
                        OutputWriter writer = new OutputWriter(referenceData, aggregator);
                        String rateOutput = new File(targetDir, "Consolidated_Rate_Report.xlsx").getAbsolutePath();
                        writer.write(rateOutput);
                        log("  > Rate Report created: " + rateOutput);
                    } else {
                        log("  - Rate Warning: No valid data found for Rate module.");
                    }
                } catch (Exception e) {
                    log("  ! Rate Module Failed: " + e.getMessage());
                    e.printStackTrace();
                }

                // 2. EXT CODE MODULE
                try {
                    log("\n[2/3] Running Forecast By ExtCode...");
                    ExcelReader reader = new ExcelReader();
                    ServiceTeamParser parser = new ServiceTeamParser();
                    forecast.by.extcode.service.ExcelWriter writer = new forecast.by.extcode.service.ExcelWriter();

                    List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
                    for (File f : inputs) {
                        try {
                            rawItems.addAll(reader.extractRawServiceTeams(f));
                        } catch (Exception e) {
                            log("  - ExtCode Warning: Failed to process " + f.getName());
                        }
                    }

                    if (!rawItems.isEmpty()) {
                        List<String> labels = new ArrayList<>();
                        for (ExcelReader.ServiceTeamRaw raw : rawItems) labels.add(raw.getLabel());
                        List<ServiceTeam> parsed = parser.parse(labels);

                        for (int i = 0; i < parsed.size(); i++) {
                            parsed.get(i).setCost(rawItems.get(i).getCost() == null ? "" : String.valueOf(rawItems.get(i).getCost()));
                            parsed.get(i).setStyle(rawItems.get(i).getCost() == null ? null : rawItems.get(i).getStyle());
                        }

                        writer.write(parsed, targetDir);
                        log("  > ExtCode Report created in: " + targetDir.getAbsolutePath());
                    } else {
                        log("  - ExtCode Warning: No valid data found.");
                    }
                } catch (Exception e) {
                    log("  ! ExtCode Module Failed: " + e.getMessage());
                    e.printStackTrace();
                }

                // 3. MONTH MODULE
                try {
                    log("\n[3/3] Running Forecast By Month...");
                    for (File f : inputs) {
                        try {
                            log("  - Processing " + f.getName() + "...");
                            ExecuteService.executeScript(f.getAbsolutePath(), targetDir.getAbsolutePath(), months);
                        } catch (Exception e) {
                            log("  - Month Warning: Failed to process " + f.getName() + ": " + e.getMessage());
                        }
                    }
                    log("  > Month processing finished.");
                } catch (Exception e) {
                    log("  ! Month Module Critical Error: " + e.getMessage());
                }

                log("\n=== EXECUTION COMPLETED ===");
                JOptionPane.showMessageDialog(this, "All processes finished. Check logs for details.");
            }).start();
        }
    }

    // ==========================================
    // 1. Rate Panel
    // ==========================================
    static class RatePanel extends JPanel {
        private final DefaultListModel<File> inputFilesModel = new DefaultListModel<>();
        private final JList<File> inputFilesList = new JList<>(inputFilesModel);
        private final JTextField outputField = new JTextField();
        private final JTextArea logArea = new JTextArea();

        public RatePanel() {
            setLayout(new BorderLayout(15, 15));
            setBorder(new EmptyBorder(20, 20, 20, 20));
            setBackground(BG_COLOR);

            // Top: Configuration
            JPanel configPanel = new JPanel(new GridBagLayout());
            configPanel.setBackground(Color.WHITE);
            configPanel.setBorder(createTitledBorder("Configuration"));
            
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 10, 10, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            // Row 0: Reference Data Note
            gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 2;
            JLabel refLabel = new JLabel("Reference Data source: src/main/resources/Data.xlsx");
            refLabel.setForeground(Color.GRAY);
            refLabel.setIcon(UIManager.getIcon("FileView.fileIcon"));
            configPanel.add(refLabel, gbc);

            // Row 1: Input Files
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 1; gbc.weightx = 0;
            configPanel.add(new JLabel("Input Files:"), gbc);
            
            gbc.gridx = 1; gbc.gridy = 1; gbc.weightx = 1.0;
            JButton selectInputsBtn = createStyledButton("Select Excel Files");
            selectInputsBtn.addActionListener(e -> selectInputs());
            configPanel.add(selectInputsBtn, gbc);

            // Row 2: Input List Scroll
            gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 0.5;
            gbc.ipady = 50; // Minimum height
            JScrollPane listScroll = new JScrollPane(inputFilesList);
            listScroll.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
            configPanel.add(listScroll, gbc);
            gbc.ipady = 0; // Reset

            // Row 3: Output File
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weighty = 0;
            gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 1;
            configPanel.add(new JLabel("Output File:"), gbc);

            gbc.gridx = 1; gbc.gridy = 3;
            JPanel outputPanel = new JPanel(new BorderLayout(10, 0));
            outputPanel.setOpaque(false);
            outputField.setEditable(false);
            outputPanel.add(outputField, BorderLayout.CENTER);
            JButton selectOutputBtn = createStyledButton("Browse...");
            selectOutputBtn.addActionListener(e -> selectOutput());
            outputPanel.add(selectOutputBtn, BorderLayout.EAST);
            configPanel.add(outputPanel, gbc);

            // Row 4: Run Button
            gbc.gridx = 0; gbc.gridy = 4; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.NONE;
            gbc.anchor = GridBagConstraints.CENTER;
            JButton runBtn = createStyledButton("Generate Consolidated Report");
            runBtn.setBackground(new Color(39, 174, 96)); // Green
            runBtn.addActionListener(e -> runProcess());
            configPanel.add(runBtn, gbc);

            add(configPanel, BorderLayout.NORTH);

            // Center: Logs
            logArea.setEditable(false);
            JScrollPane logScroll = new JScrollPane(logArea);
            logScroll.setBorder(createTitledBorder("Logs / Console Output"));
            add(logScroll, BorderLayout.CENTER);
        }

        private void selectInputs() {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
            if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                for (File f : chooser.getSelectedFiles()) {
                    if (!inputFilesModel.contains(f)) {
                        inputFilesModel.addElement(f);
                    }
                }
            }
        }

        private void selectOutput() {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
            if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                String path = chooser.getSelectedFile().getAbsolutePath();
                if (!path.endsWith(".xlsx")) path += ".xlsx";
                outputField.setText(path);
            }
        }

        private void runProcess() {
            if (inputFilesModel.isEmpty()) {
                log("Error: No input files selected.");
                return;
            }
            if (outputField.getText().isEmpty()) {
                log("Error: No output file selected.");
                return;
            }

            PrintStream originalOut = System.out;
            PrintStream captureStream = new PrintStream(new OutputStream() {
                @Override
                public void write(int b) {
                    logArea.append(String.valueOf((char) b));
                    logArea.setCaretPosition(logArea.getDocument().getLength());
                }
            });
            System.setOut(captureStream);

            new Thread(() -> {
                try {
                    log("Starting process...");
                    
                    ReferenceData referenceData = new ReferenceData();
                    String dataPath = "C:\\Users\\Sanae\\Desktop\\Task_java_excel\\ctsInvoicing\\src\\main\\resources\\Data.xlsx";
                    File dataFile = new File(dataPath);
                    if (!dataFile.exists()) {
                         dataPath = "src/main/resources/Data.xlsx";
                    }
                    log("Loading Reference Data from: " + dataPath);
                    referenceData.load(dataPath);

                    GroupAggregator aggregator = new GroupAggregator();
                    InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
                    InputFilesReader filesReader = new InputFilesReader(rowProcessor, aggregator);

                    for (int i = 0; i < inputFilesModel.size(); i++) {
                        File f = inputFilesModel.get(i);
                        log("Processing: " + f.getName());
                        filesReader.processFile(f.getAbsolutePath());
                    }
                    
                    // Print raw data grouped by user
                    rowProcessor.printRawRates();

                    OutputWriter writer = new OutputWriter(referenceData, aggregator);
                    writer.write(outputField.getText());

                    log("Success! Output written to: " + outputField.getText());
                    JOptionPane.showMessageDialog(this, "Process Completed Successfully!");

                } catch (Exception e) {
                    log("Error: " + e.getMessage());
                    e.printStackTrace(captureStream);
                    JOptionPane.showMessageDialog(this, "Error: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                } finally {
                    System.setOut(originalOut);
                }
            }).start();
        }

        private void log(String msg) {
            logArea.append(msg + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        }
    }

    // ==========================================
    // 2. Month Panel
    // ==========================================
    static class MonthPanel extends JPanel {
        private final JTextField excelFilePathField = new JTextField();
        private final JTextField targetPathField = new JTextField();
        private final JCheckBox saveCheckbox = new JCheckBox("Save the target for next use");
        private final JSpinner monthsSpinner = new JSpinner(new SpinnerNumberModel(3, 1, 12, 1));

        public MonthPanel() {
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(30, 30, 30, 30));
            setBackground(BG_COLOR);
            
            JPanel formPanel = new JPanel(new GridBagLayout());
            formPanel.setBackground(Color.WHITE);
            formPanel.setBorder(createTitledBorder("Monthly Forecast Settings"));
            
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(15, 15, 15, 15);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            
            // Row 0: Excel File
            gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0;
            formPanel.add(new JLabel("Excel File:"), gbc);

            gbc.gridx = 1; gbc.gridy = 0; gbc.weightx = 1.0;
            JPanel filePanel = new JPanel(new BorderLayout(10, 0));
            filePanel.setOpaque(false);
            excelFilePathField.setEditable(false);
            filePanel.add(excelFilePathField, BorderLayout.CENTER);
            JButton selectExcelBtn = createStyledButton("Browse...");
            selectExcelBtn.addActionListener(e -> selectExcelFile());
            filePanel.add(selectExcelBtn, BorderLayout.EAST);
            formPanel.add(filePanel, gbc);

            // Row 1: Target Directory
            gbc.gridx = 0; gbc.gridy = 1; gbc.weightx = 0;
            formPanel.add(new JLabel("Target Directory:"), gbc);

            gbc.gridx = 1; gbc.gridy = 1; gbc.weightx = 1.0;
            JPanel dirPanel = new JPanel(new BorderLayout(10, 0));
            dirPanel.setOpaque(false);
            targetPathField.setEditable(false);
            dirPanel.add(targetPathField, BorderLayout.CENTER);
            JButton selectTargetBtn = createStyledButton("Browse...");
            selectTargetBtn.addActionListener(e -> selectTargetDir());
            dirPanel.add(selectTargetBtn, BorderLayout.EAST);
            formPanel.add(dirPanel, gbc);

            // Row 2: Months Selection
            gbc.gridx = 0; gbc.gridy = 2; gbc.weightx = 0;
            formPanel.add(new JLabel("Number of Months:"), gbc);

            gbc.gridx = 1; gbc.gridy = 2; gbc.weightx = 1.0;
            // Style the spinner a bit
            JComponent editor = monthsSpinner.getEditor();
            if (editor instanceof JSpinner.DefaultEditor) {
                ((JSpinner.DefaultEditor)editor).getTextField().setColumns(5);
            }
            JPanel spinnerPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
            spinnerPanel.setOpaque(false);
            spinnerPanel.add(monthsSpinner);
            formPanel.add(spinnerPanel, gbc);

            // Row 3: Checkbox
            gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 2;
            loadSavedPath();
            saveCheckbox.setOpaque(false);
            saveCheckbox.addActionListener(e -> {
                if (!saveCheckbox.isSelected()) {
                    targetPathField.setText("");
                    savePreference("");
                }
            });
            formPanel.add(saveCheckbox, gbc);

            // Row 4: Execute
            gbc.gridx = 0; gbc.gridy = 4; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.NONE;
            gbc.anchor = GridBagConstraints.CENTER;
            JButton executeBtn = createStyledButton("Execute Script");
            executeBtn.setBackground(new Color(39, 174, 96));
            executeBtn.addActionListener(e -> execute());
            formPanel.add(executeBtn, gbc);

            // Wrapper to keep form centered at top
            JPanel wrapper = new JPanel(new BorderLayout());
            wrapper.setBackground(BG_COLOR);
            wrapper.add(formPanel, BorderLayout.NORTH);
            
            add(wrapper, BorderLayout.CENTER);
        }

        private void selectExcelFile() {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
            if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                excelFilePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
            }
        }

        private void selectTargetDir() {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                targetPathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
            }
        }

        private void execute() {
            String excelPath = excelFilePathField.getText();
            String targetPath = targetPathField.getText();

            if (excelPath.isEmpty() || targetPath.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Please select both Excel file and Target directory.", "Warning", JOptionPane.WARNING_MESSAGE);
                return;
            }

            if (saveCheckbox.isSelected()) {
                savePreference(targetPath);
            }

            int months = (Integer) monthsSpinner.getValue();

            new Thread(() -> {
                try {
                    ExecuteService.executeScript(excelPath, targetPath, months);
                    JOptionPane.showMessageDialog(this, "Execution Completed Successfully!");
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(this, "Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    ex.printStackTrace();
                }
            }).start();
        }

        private void loadSavedPath() {
            File targetFile = new File("GeneratedInvoicing.txt");
            if (targetFile.exists()) {
                try (BufferedReader reader = new BufferedReader(new FileReader(targetFile))) {
                    String savedPath = reader.readLine();
                    if (savedPath != null && !savedPath.isEmpty()) {
                        targetPathField.setText(savedPath);
                        saveCheckbox.setSelected(true);
                    }
                } catch (IOException e) {
                    // ignore
                }
            }
        }

        private void savePreference(String content) {
            try (BufferedWriter writer = new BufferedWriter(new FileWriter("GeneratedInvoicing.txt"))) {
                writer.write(content);
            } catch (IOException ex) {
                // ignore
            }
        }
    }

    // ==========================================
    // 3. ExtCode Panel
    // ==========================================
    static class ExtCodePanel extends JPanel {
        private final DefaultListModel<File> inputFilesModel = new DefaultListModel<>();
        private final JList<File> inputFilesList = new JList<>(inputFilesModel);
        private final JTextField targetDirField = new JTextField();

        public ExtCodePanel() {
            setLayout(new BorderLayout(15, 15));
            setBorder(new EmptyBorder(20, 20, 20, 20));
            setBackground(BG_COLOR);

            JPanel mainPanel = new JPanel(new GridBagLayout());
            mainPanel.setBackground(Color.WHITE);
            mainPanel.setBorder(createTitledBorder("External Code Export"));
            
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 10, 10, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            // Input Files
            gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0;
            mainPanel.add(new JLabel("Input Excel Files:"), gbc);
            
            gbc.gridx = 1; gbc.gridy = 0; gbc.weightx = 1.0;
            JButton selectInputsBtn = createStyledButton("Select Files");
            selectInputsBtn.addActionListener(e -> selectInputs());
            mainPanel.add(selectInputsBtn, gbc);

            // List
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 0.5;
            gbc.ipady = 50;
            JScrollPane listScroll = new JScrollPane(inputFilesList);
            listScroll.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
            mainPanel.add(listScroll, gbc);
            gbc.ipady = 0;

            // Target Dir
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weighty = 0;
            gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 1;
            mainPanel.add(new JLabel("Output Folder:"), gbc);

            gbc.gridx = 1; gbc.gridy = 2;
            JPanel dirPanel = new JPanel(new BorderLayout(10, 0));
            dirPanel.setOpaque(false);
            targetDirField.setEditable(false);
            dirPanel.add(targetDirField, BorderLayout.CENTER);
            JButton selectTargetBtn = createStyledButton("Browse...");
            selectTargetBtn.addActionListener(e -> selectTarget());
            dirPanel.add(selectTargetBtn, BorderLayout.EAST);
            mainPanel.add(dirPanel, gbc);

            // Run
            gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.NONE;
            gbc.anchor = GridBagConstraints.CENTER;
            JButton runBtn = createStyledButton("Export Excel");
            runBtn.setBackground(new Color(39, 174, 96));
            runBtn.addActionListener(e -> runProcess());
            mainPanel.add(runBtn, gbc);

            add(mainPanel, BorderLayout.CENTER);
        }

        private void selectInputs() {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                for (File f : chooser.getSelectedFiles()) {
                    if (!inputFilesModel.contains(f)) {
                        inputFilesModel.addElement(f);
                    }
                }
            }
        }

        private void selectTarget() {
            JFileChooser chooser = new JFileChooser();
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                targetDirField.setText(chooser.getSelectedFile().getAbsolutePath());
            }
        }

        private void runProcess() {
            if (inputFilesModel.isEmpty() || targetDirField.getText().isEmpty()) {
                JOptionPane.showMessageDialog(this, "Please select input files and output folder.", "Warning", JOptionPane.WARNING_MESSAGE);
                return;
            }

            new Thread(() -> {
                try {
                    ExcelReader reader = new ExcelReader();
                    ServiceTeamParser parser = new ServiceTeamParser();
                    forecast.by.extcode.service.ExcelWriter writer = new forecast.by.extcode.service.ExcelWriter();

                    List<File> files = new ArrayList<>();
                    for(int i=0; i<inputFilesModel.size(); i++) files.add(inputFilesModel.get(i));

                    List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
                    for (File f : files) {
                        rawItems.addAll(reader.extractRawServiceTeams(f));
                    }

                    List<String> labels = new ArrayList<>();
                    for (ExcelReader.ServiceTeamRaw raw : rawItems) {
                        labels.add(raw.getLabel());
                    }

                    List<ServiceTeam> parsed = parser.parse(labels);

                    // Attach cost + style
                    for (int i = 0; i < parsed.size(); i++) {
                        parsed.get(i).setCost(
                                rawItems.get(i).getCost() == null ? "" : String.valueOf(rawItems.get(i).getCost())
                        );
                        parsed.get(i).setStyle(
                                rawItems.get(i).getCost() == null ? null : rawItems.get(i).getStyle()
                        );
                    }

                    File target = new File(targetDirField.getText());
                    writer.write(parsed, target);

                    JOptionPane.showMessageDialog(this, "Excel exported successfully!");

                } catch (Exception e) {
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(this, "Error: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }).start();
        }
    }
}
