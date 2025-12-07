package forecast.by.rate.service;
// InputFilesReader.java
import java.io.FileInputStream;
import java.io.IOException;

import forecast.by.rate.util.GroupAggregator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Map;

public class InputFilesReader {
    private final InputRowProcessor rowProcessor;
    private final GroupAggregator aggregator;

    public InputFilesReader(InputRowProcessor rowProcessor, GroupAggregator aggregator) {
        this.rowProcessor = rowProcessor;
        this.aggregator = aggregator;
    }

    public void processFile(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet = wb.getSheetAt(0); // Assuming data is in the first sheet
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header
                Map.Entry<String, Double> result = rowProcessor.processRow(row);
                if (result != null) {
                    aggregator.addHoras(result.getKey(), result.getValue());
                }
            }
        }
    }
}