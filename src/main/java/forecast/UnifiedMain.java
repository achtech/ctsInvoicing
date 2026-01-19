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
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class UnifiedMain extends JFrame {

    public UnifiedMain() {
        setTitle("CTS Invoicing Unified Tool");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(900, 650);
        setLocationRelativeTo(null);

        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.addTab("Forecast By Rate", new RatePanel());
        tabbedPane.addTab("Forecast By Month", new MonthPanel());
        tabbedPane.addTab("Forecast By ExtCode", new ExtCodePanel());

        add(tabbedPane);
    }

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }
        SwingUtilities.invokeLater(() -> new UnifiedMain().setVisible(true));
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
            setLayout(new BorderLayout(10, 10));
            setBorder(new EmptyBorder(10, 10, 10, 10));

            // Top: Configuration
            JPanel topPanel = new JPanel(new GridBagLayout());
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(5, 5, 5, 5);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            // Row 0: Reference Data Note
            gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 2;
            JLabel refLabel = new JLabel("Reference Data: src/main/resources/Data.xlsx (Hardcoded)");
            refLabel.setForeground(Color.GRAY);
            topPanel.add(refLabel, gbc);

            // Row 1: Input Files
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 1;
            topPanel.add(new JLabel("Input Files:"), gbc);
            
            gbc.gridx = 1; gbc.gridy = 1;
            JButton selectInputsBtn = new JButton("Select Excel Files");
            selectInputsBtn.addActionListener(e -> selectInputs());
            topPanel.add(selectInputsBtn, gbc);

            // Row 2: Input List Scroll
            gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 0.3;
            topPanel.add(new JScrollPane(inputFilesList), gbc);

            // Row 3: Output File
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weighty = 0;
            gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 1;
            topPanel.add(new JLabel("Output File:"), gbc);

            gbc.gridx = 1; gbc.gridy = 3;
            JPanel outputPanel = new JPanel(new BorderLayout(5, 0));
            outputField.setEditable(false);
            outputPanel.add(outputField, BorderLayout.CENTER);
            JButton selectOutputBtn = new JButton("Select Output");
            selectOutputBtn.addActionListener(e -> selectOutput());
            outputPanel.add(selectOutputBtn, BorderLayout.EAST);
            topPanel.add(outputPanel, gbc);

            // Row 4: Run Button
            gbc.gridx = 0; gbc.gridy = 4; gbc.gridwidth = 2;
            JButton runBtn = new JButton("Generate Consolidated Report");
            runBtn.setFont(runBtn.getFont().deriveFont(Font.BOLD, 14f));
            runBtn.addActionListener(e -> runProcess());
            topPanel.add(runBtn, gbc);

            add(topPanel, BorderLayout.NORTH);

            // Center: Logs
            logArea.setEditable(false);
            logArea.setFont(new Font("Monospaced", Font.PLAIN, 12));
            JScrollPane logScroll = new JScrollPane(logArea);
            logScroll.setBorder(BorderFactory.createTitledBorder("Logs / Console Output"));
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

            // Redirect stdout to logArea for this thread? 
            // Better to just print and also append manually, 
            // but the underlying classes use System.out.println.
            // For now, we'll capture standard out.
            PrintStream originalOut = System.out;
            PrintStream captureStream = new PrintStream(new OutputStream() {
                @Override
                public void write(int b) {
                    logArea.append(String.valueOf((char) b));
                    // Also scroll to bottom
                    logArea.setCaretPosition(logArea.getDocument().getLength());
                }
            });
            System.setOut(captureStream);

            new Thread(() -> {
                try {
                    log("Starting process...");
                    
                    ReferenceData referenceData = new ReferenceData();
                    // Try absolute path first (as in original code), then relative
                    String dataPath = "C:\\Users\\Sanae\\Desktop\\Task_java_excel\\ctsInvoicing\\src\\main\\resources\\Data.xlsx";
                    File dataFile = new File(dataPath);
                    if (!dataFile.exists()) {
                         // Fallback to simpler path
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

                    OutputWriter writer = new OutputWriter(referenceData, aggregator);
                    writer.write(outputField.getText());

                    log("Success! Output written to: " + outputField.getText());
                    JOptionPane.showMessageDialog(this, "Process Completed Successfully!");

                } catch (Exception e) {
                    log("Error: " + e.getMessage());
                    e.printStackTrace(captureStream); // Print stack trace to log
                    JOptionPane.showMessageDialog(this, "Error: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                } finally {
                    System.setOut(originalOut); // Restore
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
            setLayout(new BorderLayout());
            
            // Replicating logic from forecast.by.month.Main
            // We use a simpler form layout instead of absolute positioning for better resizing
            JPanel formPanel = new JPanel(new GridBagLayout());
            formPanel.setBorder(new EmptyBorder(20, 20, 20, 20));
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 10, 10, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            
            // Row 0: Excel File
            gbc.gridx = 0; gbc.gridy = 0;
            JButton selectExcelBtn = new JButton("Select Excel File");
            selectExcelBtn.addActionListener(e -> selectExcelFile());
            formPanel.add(selectExcelBtn, gbc);

            gbc.gridx = 1; gbc.gridy = 0; gbc.weightx = 1.0;
            excelFilePathField.setEditable(false);
            formPanel.add(excelFilePathField, gbc);

            // Row 1: Target Directory
            gbc.gridx = 0; gbc.gridy = 1; gbc.weightx = 0;
            JButton selectTargetBtn = new JButton("Select Target Dir");
            selectTargetBtn.addActionListener(e -> selectTargetDir());
            formPanel.add(selectTargetBtn, gbc);

            gbc.gridx = 1; gbc.gridy = 1; gbc.weightx = 1.0;
            targetPathField.setEditable(false);
            formPanel.add(targetPathField, gbc);

            // Row 2: Months Selection
            gbc.gridx = 0; gbc.gridy = 2; gbc.weightx = 0;
            formPanel.add(new JLabel("Number of Months:"), gbc);

            gbc.gridx = 1; gbc.gridy = 2; gbc.weightx = 1.0;
            formPanel.add(monthsSpinner, gbc);

            // Row 3: Checkbox
            gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 2;
            // Load saved preference
            loadSavedPath();
            saveCheckbox.addActionListener(e -> {
                if (!saveCheckbox.isSelected()) {
                    targetPathField.setText("");
                    savePreference("");
                }
            });
            formPanel.add(saveCheckbox, gbc);

            // Row 4: Execute
            gbc.gridx = 0; gbc.gridy = 4; gbc.gridwidth = 2;
            JButton executeBtn = new JButton("Execute");
            executeBtn.setFont(executeBtn.getFont().deriveFont(Font.BOLD, 14f));
            executeBtn.addActionListener(e -> execute());
            formPanel.add(executeBtn, gbc);

            add(formPanel, BorderLayout.CENTER);

            // Add background image hint if needed, but standard UI is safer
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
            setLayout(new BorderLayout(10, 10));
            setBorder(new EmptyBorder(10, 10, 10, 10));

            JPanel topPanel = new JPanel(new GridBagLayout());
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(5, 5, 5, 5);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            // Input Files
            gbc.gridx = 0; gbc.gridy = 0;
            topPanel.add(new JLabel("Input Excel Files:"), gbc);
            
            gbc.gridx = 1; gbc.gridy = 0;
            JButton selectInputsBtn = new JButton("Select Files");
            selectInputsBtn.addActionListener(e -> selectInputs());
            topPanel.add(selectInputsBtn, gbc);

            // List
            gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 0.5;
            topPanel.add(new JScrollPane(inputFilesList), gbc);

            // Target Dir
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.weighty = 0;
            gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 1;
            topPanel.add(new JLabel("Output Folder:"), gbc);

            gbc.gridx = 1; gbc.gridy = 2;
            JPanel targetPanel = new JPanel(new BorderLayout(5, 0));
            targetDirField.setEditable(false);
            targetPanel.add(targetDirField, BorderLayout.CENTER);
            JButton selectTargetBtn = new JButton("Select Folder");
            selectTargetBtn.addActionListener(e -> selectTarget());
            targetPanel.add(selectTargetBtn, BorderLayout.EAST);
            topPanel.add(targetPanel, gbc);

            // Run
            gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 2;
            JButton runBtn = new JButton("Export Excel");
            runBtn.setFont(runBtn.getFont().deriveFont(Font.BOLD, 14f));
            runBtn.addActionListener(e -> runProcess());
            topPanel.add(runBtn, gbc);

            add(topPanel, BorderLayout.CENTER);
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
