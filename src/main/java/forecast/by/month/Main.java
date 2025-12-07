package forecast.by.month;

import forecast.by.month.service.ExecuteService;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

public class Main extends JFrame {
    private final JTextField excelFilePathField;
    private JTextField targetPathField;
    private JCheckBox saveCheckbox;
    int xStart1=35,xStart3=350,yStart1 = 150, xStart2=xStart1+135,yStart2 = yStart1 + 60 ,yStart3 = yStart2+60,yStart4 = yStart3+40;
    public Main() {
        // Set up the main frame
        setTitle("File Selection Interface");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(520, 520);
        setLocationRelativeTo(null);
        setLayout(new BorderLayout()); // Use BorderLayout for the frame

        // Create a JLabel with the background image
        JLabel background = new JLabel(new ImageIcon("Nttdata_back.png"));
        background.setBounds(0, 0, 520, 520); // Match frame size
        background.setLayout(null); // Use null layout for precise positioning on background

        // Excel file selection components
        JButton selectExcelButton = new JButton("Select Excel File");
        JButton selectTargetButton = new JButton("Select Target");
        JButton executeButton = new JButton("Execute");
        excelFilePathField = new JTextField(30);

        // Load target path from file if it exists and is not empty
        File targetFile = new File("GeneratedInvoicing.txt");
        if (targetFile.exists()) {
            try (BufferedReader reader = new BufferedReader(new FileReader(targetFile))) {
                String savedPath = reader.readLine();
                if (savedPath != null && !savedPath.isEmpty()) {
                    targetPathField = new JTextField(savedPath, 30);
                    saveCheckbox = new JCheckBox("Save the target for next use");
                    saveCheckbox.setForeground(Color.WHITE);
                    saveCheckbox.setSelected(true);
                } else {
                    targetPathField = new JTextField(30);
                    saveCheckbox = new JCheckBox("Save the target for next use");
                }
            } catch (IOException e) {
                targetPathField = new JTextField(30);
                saveCheckbox = new JCheckBox("Save the target for next use");
            }
        } else {
            targetPathField = new JTextField(30);
            saveCheckbox = new JCheckBox("Save the target for next use");
        }
        targetPathField.setEditable(false);

        selectExcelButton.setBackground(Color.WHITE); // #2B2490
        excelFilePathField.setEditable(false);

        selectExcelButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
                int result = fileChooser.showOpenDialog(Main.this);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    excelFilePathField.setText(selectedFile.getAbsolutePath());
                }
            }
        });

        // Target selection components
        selectTargetButton.setBackground(Color.white); // #2B2490

        selectTargetButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int result = fileChooser.showOpenDialog(Main.this);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedDir = fileChooser.getSelectedFile();
                    targetPathField.setText(selectedDir.getAbsolutePath());
                }
            }
        });

        // Execute button
        executeButton.setBackground(Color.WHITE);
        executeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (saveCheckbox.isSelected()) {
                    try (BufferedWriter writer = new BufferedWriter(new FileWriter("GeneratedInvoicing.txt"))) {
                        writer.write(targetPathField.getText());
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(Main.this, "Failed to save target path.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
                ExecuteService.executeScript(excelFilePathField.getText(), targetPathField.getText());
            }
        });

        // Save checkbox with action listener to clear text and file when unchecked
        saveCheckbox.setForeground(Color.WHITE);
        saveCheckbox.setOpaque(false); // Make checkbox background transparent
        saveCheckbox.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!saveCheckbox.isSelected()) {
                    targetPathField.setText("");
                    try (BufferedWriter writer = new BufferedWriter(new FileWriter("GeneratedInvoicing.txt"))) {
                        writer.write("");
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(Main.this, "Failed to clear target path file.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        });

        excelFilePathField.setBounds(xStart2, yStart1, 280, 30); // x, y, width, height
        selectExcelButton.setBounds(xStart1, yStart1, 130, 30); // x, y, width, height
        targetPathField.setBounds(xStart2, yStart2, 280, 30); // x, y, width, height
        selectTargetButton.setBounds(xStart1, yStart2, 130, 30); // x, y, width, height
        saveCheckbox.setBounds(xStart1, yStart3, 200, 20); // x, y, width, height
        executeButton.setBounds(xStart3, yStart3, 100, 30); // x, y, width, height

        // Add components to the background label
        background.add(selectExcelButton);
        background.add(excelFilePathField);
        background.add(selectTargetButton);
        background.add(targetPathField);
        background.add(executeButton);
        background.add(saveCheckbox);

        // Add the background label to the frame
        add(background, BorderLayout.CENTER);

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            Main frame = new Main();
            frame.setVisible(true);
        });
    }
}