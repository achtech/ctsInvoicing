package forecast.by.rate;

// Main.java
import forecast.by.rate.service.InputFilesReader;
import forecast.by.rate.service.InputRowProcessor;
import forecast.by.rate.service.OutputWriter;
import forecast.by.rate.util.GroupAggregator;
import forecast.by.rate.util.ReferenceData;

import java.io.File;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Main {
    public static void main(String[] args) {
        // UI for selecting input files
        JFileChooser inputChooser = new JFileChooser();
        inputChooser.setMultiSelectionEnabled(true);
        inputChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        int inputOption = inputChooser.showOpenDialog(null);
        if (inputOption != JFileChooser.APPROVE_OPTION) {
            System.out.println("No input files selected.");
            return;
        }
        File[] inputFiles = inputChooser.getSelectedFiles();

        // UI for selecting output file
        JFileChooser outputChooser = new JFileChooser();
        outputChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        int outputOption = outputChooser.showSaveDialog(null);
        if (outputOption != JFileChooser.APPROVE_OPTION) {
            System.out.println("No output file selected.");
            return;
        }
        String outputPath = outputChooser.getSelectedFile().getAbsolutePath();
        if (!outputPath.endsWith(".xlsx")) {
            outputPath += ".xlsx";
        }

        try {
            // Load reference data (adjust path as needed; assumes 'resources/Data.xlsx')
            ReferenceData referenceData = new ReferenceData();
//            referenceData.load("Data.xlsx");
//            referenceData.load( "src/main/resources/Data.xlsx" );
            referenceData.load( "C:\\Users\\Sanae\\Desktop\\Task_java_excel\\ctsInvoicing\\src\\main\\resources\\Data.xlsx" );

            GroupAggregator aggregator = new GroupAggregator();
            InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
            InputFilesReader filesReader = new InputFilesReader(rowProcessor, aggregator);

            for (File inputFile : inputFiles) {
                filesReader.processFile(inputFile.getAbsolutePath());
            }

            OutputWriter writer = new OutputWriter(referenceData, aggregator);
            writer.write(outputPath);

            System.out.println("Output file generated successfully at: " + outputPath);
        } catch (IOException e) {
            System.err.println("Error processing files: " + e.getMessage());
            e.printStackTrace();
        }
    }
}