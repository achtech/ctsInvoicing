package invoicing.service.rate;

import invoicing.Helper.GroupAggregator;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.Locale;

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
            Sheet sheet = findSheet(wb, null, true);
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                InputRowProcessor.RowData result = rowProcessor.processRow(row, evaluator);
                if (result != null) {
                    aggregator.add(result.getGroupId(), "user", result.getHours(), result.getCost());
                }
            }
        }
    }

    public boolean processFile(String filePath, String monthSpanish) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet = findSheet(wb, monthSpanish, false);
            if (sheet == null) {
                return false;
            }

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                InputRowProcessor.RowData result = rowProcessor.processRow(row, evaluator);
                if (result != null) {
                    aggregator.add(result.getGroupId(), "user", result.getHours(), result.getCost());
                }
            }
            return true;
        }
    }

    private Sheet findSheet(Workbook wb, String monthSpanish, boolean fallbackAnyFacturacion) {
        String normalizedTargetMonth = normalize(monthSpanish);
        Sheet firstFacturacion = null;

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            String normalizedName = normalize(s.getSheetName());
            if (!normalizedName.contains("facturacion")) {
                continue;
            }

            if (firstFacturacion == null) {
                firstFacturacion = s;
            }

            if (normalizedTargetMonth != null && !normalizedTargetMonth.isBlank() && normalizedName.contains(normalizedTargetMonth)) {
                return s;
            }
        }

        if (normalizedTargetMonth != null && !normalizedTargetMonth.isBlank()) {
            return null;
        }

        if (firstFacturacion != null) {
            return firstFacturacion;
        }

        return fallbackAnyFacturacion ? wb.getSheetAt(0) : null;
    }

    private String normalize(String value) {
        if (value == null) {
            return null;
        }
        return Normalizer.normalize(value, Normalizer.Form.NFD)
                .replaceAll("\\p{M}+", "")
                .toLowerCase(Locale.ROOT);
    }
}