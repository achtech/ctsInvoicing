package invoicing.view;

import invoicing.service.UnifiedExecutionService;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class InvoicingDashboard extends JFrame {

    private static final Color PRIMARY_COLOR = new Color(41, 128, 185);
    private static final Color BG_COLOR = new Color(245, 245, 245);
    private static final Color TEXT_COLOR = new Color(44, 62, 80);
    private static final Font MAIN_FONT = new Font("Segoe UI", Font.PLAIN, 13);
    private static final Font HEADER_FONT = new Font("Segoe UI", Font.BOLD, 14);

    public InvoicingDashboard() {
        setTitle("CTS Invoicing Dashboard");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1100, 820);
        setMinimumSize(new Dimension(900, 700));
        setLocationRelativeTo(null);
        setBackground(BG_COLOR);
        add(new AllInOnePanel());
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

    static class AllInOnePanel extends JPanel {

        private final DefaultListModel<File> inputFilesModel = new DefaultListModel<>();
        private final JList<File> inputFilesList = new JList<>(inputFilesModel);
        private final JTextField targetDirField = new JTextField();
        private final JSpinner monthsSpinner = new JSpinner(new SpinnerNumberModel(3, 1, 12, 1));
        private final JCheckBox monthsToggle = new JCheckBox("Enable", true);
        private final JLabel inputErrorLabel = new JLabel(" ");
        private final JLabel outputErrorLabel = new JLabel(" ");
        private final JTextArea logArea = new JTextArea();

        private final JProgressBar progressBar = new JProgressBar(0, 3);
        private final JLabel statusLabel = new JLabel("Ready");

        private final JButton runBtn = createStyledButton("RUN ALL PROCESSES");
        private final JButton openOutputBtn = createStyledButton("OPEN MAIN OUTPUT FOLDER");

        private File lastMainOutputFolder;

        private static final String DEFAULT_HISTORY_PATH = "src/main/resources/history.csv";
        private static final String LOCAL_HISTORY_PATH = "history.csv";

        private String getEffectiveHistoryPath() {
            if (new File(DEFAULT_HISTORY_PATH).exists()) {
                return DEFAULT_HISTORY_PATH;
            }
            return LOCAL_HISTORY_PATH;
        }

        public AllInOnePanel() {
            setLayout(new BorderLayout(10, 10));
            setBorder(new EmptyBorder(15, 15, 15, 15));
            setBackground(BG_COLOR);

            JPanel configPanel = new JPanel(new GridBagLayout());
            configPanel.setBackground(Color.WHITE);
            configPanel.setBorder(createTitledBorder("Unified Process Configuration"));

            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(6, 10, 2, 10);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.weightx = 0;
            gbc.gridwidth = 1;
            configPanel.add(new JLabel("Input Excel Files:"), gbc);

            gbc.gridx = 0;
            gbc.gridy = 1;
            gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.BOTH;
            gbc.weighty = 1.0;
            inputFilesList.setFixedCellHeight(22);
            JScrollPane listScroll = new JScrollPane(inputFilesList);
            listScroll.setPreferredSize(new Dimension(0, 90));
            listScroll.setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
            configPanel.add(listScroll, gbc);
            gbc.weighty = 0;

            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.gridx = 0;
            gbc.gridy = 2;
            gbc.gridwidth = 2;
            gbc.insets = new Insets(0, 10, 0, 10);
            inputErrorLabel.setForeground(new Color(192, 57, 43));
            inputErrorLabel.setFont(MAIN_FONT.deriveFont(Font.BOLD, 11f));
            configPanel.add(inputErrorLabel, gbc);
            gbc.insets = new Insets(4, 10, 4, 10);

            gbc.fill = GridBagConstraints.NONE;
            gbc.gridx = 0;
            gbc.gridy = 3;
            gbc.gridwidth = 2;
            gbc.anchor = GridBagConstraints.EAST;
            JPanel fileButtonsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
            fileButtonsPanel.setOpaque(false);
            JButton addBtn = createStyledButton("Add Files");
            JButton removeBtn = createStyledButton("Remove Selected");
            JButton clearBtn = createStyledButton("Clear All");
            addBtn.addActionListener(e -> selectInputs());
            removeBtn.addActionListener(e -> removeSelectedInputs());
            clearBtn.addActionListener(e -> clearInputs());
            fileButtonsPanel.add(addBtn);
            fileButtonsPanel.add(removeBtn);
            fileButtonsPanel.add(clearBtn);
            configPanel.add(fileButtonsPanel, gbc);
            gbc.anchor = GridBagConstraints.WEST;

            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.gridx = 0;
            gbc.gridy = 4;
            gbc.gridwidth = 1;
            gbc.weightx = 0;
            configPanel.add(new JLabel("Target Output Directory:"), gbc);

            gbc.gridx = 1;
            gbc.gridy = 4;
            gbc.weightx = 1.0;
            JPanel dirPanel = new JPanel(new BorderLayout(8, 0));
            dirPanel.setOpaque(false);
            targetDirField.setEditable(false);
            dirPanel.add(targetDirField, BorderLayout.CENTER);
            JButton selectTargetBtn = createStyledButton("Browse...");
            selectTargetBtn.addActionListener(e -> selectTarget());
            dirPanel.add(selectTargetBtn, BorderLayout.EAST);
            configPanel.add(dirPanel, gbc);

            gbc.gridx = 0;
            gbc.gridy = 5;
            gbc.gridwidth = 2;
            gbc.insets = new Insets(0, 10, 0, 10);
            outputErrorLabel.setForeground(new Color(192, 57, 43));
            outputErrorLabel.setFont(MAIN_FONT.deriveFont(Font.BOLD, 11f));
            configPanel.add(outputErrorLabel, gbc);
            gbc.insets = new Insets(4, 10, 4, 10);

            gbc.gridx = 0;
            gbc.gridy = 6;
            gbc.gridwidth = 1;
            gbc.weightx = 0;
            configPanel.add(new JLabel("Forecast Months (Month Module):"), gbc);

            gbc.gridx = 1;
            gbc.gridy = 6;
            gbc.weightx = 1.0;
            JComponent editor = monthsSpinner.getEditor();
            if (editor instanceof JSpinner.DefaultEditor) {
                ((JSpinner.DefaultEditor) editor).getTextField().setColumns(4);
            }
            JPanel spinnerPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
            spinnerPanel.setOpaque(false);
            monthsToggle.setOpaque(false);
            monthsToggle.setFont(MAIN_FONT);
            monthsToggle.addActionListener(e -> monthsSpinner.setEnabled(monthsToggle.isSelected()));
            spinnerPanel.add(monthsToggle);
            spinnerPanel.add(Box.createHorizontalStrut(10));
            spinnerPanel.add(monthsSpinner);
            configPanel.add(spinnerPanel, gbc);

            gbc.gridx = 0;
            gbc.gridy = 7;
            gbc.gridwidth = 2;
            gbc.fill = GridBagConstraints.NONE;
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
            progressPanel.add(progressBar, BorderLayout.CENTER);
            progressPanel.add(statusLabel, BorderLayout.SOUTH);
            progressPanel.setPreferredSize(new Dimension(0, 85));

            logArea.setEditable(false);
            logArea.setLineWrap(false);
            JScrollPane logScroll = new JScrollPane(logArea);
            logScroll.setBorder(createTitledBorder("Execution Logs"));

            JPanel bottomPanel = new JPanel(new BorderLayout(0, 8));
            bottomPanel.setOpaque(false);
            bottomPanel.add(progressPanel, BorderLayout.NORTH);
            bottomPanel.add(logScroll, BorderLayout.CENTER);

            JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT, configPanel, bottomPanel);
            split.setResizeWeight(0.50);
            split.setDividerSize(7);
            split.setBorder(null);
            split.setOpaque(false);

            add(split, BorderLayout.CENTER);

            loadHistory();
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

        private void removeSelectedInputs() {
            List<File> selected = inputFilesList.getSelectedValuesList();
            if (selected == null || selected.isEmpty()) {
                return;
            }
            for (File f : selected) {
                inputFilesModel.removeElement(f);
            }
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

        private void appendHistory(String path, int months) {
            File file = new File(getEffectiveHistoryPath());
            File parent = file.getParentFile();
            if (parent != null && !parent.exists()) {
                parent.mkdirs();
            }
            try (BufferedWriter bw = new BufferedWriter(new FileWriter(file, false))) {
                bw.write(java.time.LocalDateTime.now() + ";" + path + ";" + months);
                bw.newLine();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        private void loadHistory() {
            File file = new File(getEffectiveHistoryPath());
            if (!file.exists()) {
                return;
            }
            String lastLine = null;
            try (BufferedReader br = new BufferedReader(new FileReader(file))) {
                String line;
                while ((line = br.readLine()) != null) {
                    if (!line.trim().isEmpty()) {
                        lastLine = line;
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (lastLine == null) {
                return;
            }
            String[] parts = lastLine.split(";", -1);
            if (parts.length >= 3) {
                if (!parts[1].isEmpty()) {
                    targetDirField.setText(parts[1]);
                }
                try {
                    monthsSpinner.setValue(Integer.parseInt(parts[2]));
                } catch (NumberFormatException ignored) {
                }
            }
        }

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
            if (hasError) {
                return;
            }

            File targetDir = new File(targetDirField.getText());
            if (!targetDir.exists() || !targetDir.isDirectory()) {
                JOptionPane.showMessageDialog(this, "Invalid target directory.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            int months = (Integer) monthsSpinner.getValue();
            boolean useManual = monthsToggle.isSelected();
            appendHistory(targetDir.getAbsolutePath(), months);

            List<File> inputs = new ArrayList<>();
            for (int i = 0; i < inputFilesModel.size(); i++) {
                inputs.add(inputFilesModel.get(i));
            }

            runBtn.setEnabled(false);
            runBtn.setText("Running...");
            openOutputBtn.setEnabled(false);
            resetProgress();

            UnifiedExecutionService service = new UnifiedExecutionService();
            UnifiedExecutionService.Listener listener = new UnifiedExecutionService.Listener() {
                @Override
                public void log(String message) {
                    AllInOnePanel.this.log(message);
                }

                @Override
                public void setProgress(int value, String barLabel, String detail) {
                    AllInOnePanel.this.setProgress(value, barLabel, detail);
                }
            };

            new Thread(() -> {
                File mainOutputFolder = service.runUnified(targetDir, inputs, months, useManual, listener);

                SwingUtilities.invokeLater(() -> {
                    runBtn.setEnabled(true);
                    runBtn.setText("RUN ALL PROCESSES");
                    if (mainOutputFolder != null) {
                        lastMainOutputFolder = mainOutputFolder;
                        openOutputBtn.setEnabled(true);
                        openOutputBtn.setText("OPEN: " + mainOutputFolder.getName());
                        JOptionPane.showMessageDialog(this,
                                "All processes finished. Check logs for details.\nOutput: "
                                        + mainOutputFolder.getAbsolutePath());
                    } else {
                        JOptionPane.showMessageDialog(this,
                                "Processing finished with errors. Check logs for details.",
                                "Warning",
                                JOptionPane.WARNING_MESSAGE);
                    }
                });
            }).start();
        }

        private void openOutputFolder() {
            if (lastMainOutputFolder == null) {
                return;
            }
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
    }
}

