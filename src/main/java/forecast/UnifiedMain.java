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

        // Just add the All-in-One panel directly
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

            // Row 3: Months Spinner (Restored)
            gbc.gridx = 0; gbc.gridy = 3; gbc.weightx = 0;
            configPanel.add(new JLabel("Forecast Months (for Month Module):"), gbc);

            gbc.gridx = 1; gbc.gridy = 3; gbc.weightx = 1.0;
            // Style the spinner a bit
            JComponent editor = monthsSpinner.getEditor();
            if (editor instanceof JSpinner.DefaultEditor) {
                ((JSpinner.DefaultEditor)editor).getTextField().setColumns(5);
            }
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

            // Create Main Output Folder
            java.time.LocalDateTime now = java.time.LocalDateTime.now();
            java.time.format.DateTimeFormatter monthFormatter = java.time.format.DateTimeFormatter.ofPattern("MMM_yyyy");
            // Timestamp format updated as per user request: "your_month_day_h_min_sec" (yyyy_MM_dd_HH_mm_ss)
            java.time.format.DateTimeFormatter tsFormatter = java.time.format.DateTimeFormatter.ofPattern("yyyy_MM_dd_HH_mm_ss");
            
            String currentMonthStr = now.format(monthFormatter);
            String mainFolderName = "forecast_italy_" + currentMonthStr + "_" + now.format(tsFormatter);
            File mainOutputFolder = new File(targetDir, mainFolderName);
            if (!mainOutputFolder.exists()) {
                mainOutputFolder.mkdirs();
            }

            // Create Sub-folders
            File rateFolder = new File(mainOutputFolder, "forecast_it_rate_" + currentMonthStr);
            if (!rateFolder.exists()) rateFolder.mkdirs();

            File extFolder = new File(mainOutputFolder, "forecast_EXT_" + currentMonthStr);
            if (!extFolder.exists()) extFolder.mkdirs();

            File monthFolder = new File(mainOutputFolder, "forecast_month_" + currentMonthStr); // Optional but good for consistency
            if (!monthFolder.exists()) monthFolder.mkdirs();

            // Months value from spinner
            int months = (Integer) monthsSpinner.getValue();
            List<File> inputs = new ArrayList<>();
            for(int i=0; i<inputFilesModel.size(); i++) inputs.add(inputFilesModel.get(i));

            new Thread(() -> {
                log("=== STARTING UNIFIED EXECUTION ===");
                log("Main Output Folder: " + mainOutputFolder.getAbsolutePath());
                log("Selected Months for Month Module: " + months);
                
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
                        String rateOutput = new File(rateFolder, "Consolidated_Rate_Report.xlsx").getAbsolutePath();
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

                        writer.write(parsed, extFolder);
                        log("  > ExtCode Report created in: " + extFolder.getAbsolutePath());
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
                            log("  - Processing " + f.getName() + " with " + months + " months...");
                            ExecuteService.executeScript(f.getAbsolutePath(), monthFolder.getAbsolutePath(), months);
                        } catch (Exception e) {
                            log("  - Month Warning: Failed to process " + f.getName() + ": " + e.getMessage());
                        }
                    }
                    log("  > Month processing finished.");
                } catch (Exception e) {
                    log("  ! Month Module Critical Error: " + e.getMessage());
                }

                log("\n=== EXECUTION COMPLETED ===");
                JOptionPane.showMessageDialog(this, "All processes finished. Check logs for details.\nOutput: " + mainOutputFolder.getAbsolutePath());
            }).start();
        }


    }




}
