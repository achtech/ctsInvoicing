package invoicing.service.rate;
// InputFilesReader.java
import java.io.FileInputStream;
import java.io.IOException;

import invoicing.Helper.GroupAggregator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
            Sheet sheet = findSheet(wb); 
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header
                InputRowProcessor.RowData result = rowProcessor.processRow(row);
                if (result != null) {
                    aggregator.add(result.getGroupId(), "user", result.getHours(), result.getCost());
                }
            }
        }
    }

    private Sheet findSheet(Workbook wb) {
        String currentMonth = java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.forLanguageTag("es-ES"));
        
        Sheet match = null;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            String name = s.getSheetName();
            if (name.toLowerCase().contains("facturaci\u00F3n") && name.toLowerCase().contains(currentMonth.toLowerCase())) {
                return s;
            }
            if (match == null && name.toLowerCase().contains("facturaci\u00F3n")) {
                match = s;
            }
        }
        return match != null ? match : wb.getSheetAt(0);
    }
}
