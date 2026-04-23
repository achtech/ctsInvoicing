package invoicing.service.ext;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class ExcelReader {

    public static class ServiceTeamRaw {
        private String label;
        private Double cost;
        private CellStyle style;

        public ServiceTeamRaw(String label, Double cost, CellStyle style) {
            this.label = label;
            this.cost = cost;
            this.style = style;
        }

        public String getLabel() {
            return label;
        }

        public Double getCost() {
            return cost;
        }

        public CellStyle getStyle() {
            return style;
        }
    }

    public List<ServiceTeamRaw> extractRawServiceTeams(File file) throws Exception {
        return extractRawServiceTeams(file, null);
    }

    public List<ServiceTeamRaw> extractRawServiceTeams(File file, String monthSpanish) throws Exception {
        List<ServiceTeamRaw> result = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = findSheet(wb, monthSpanish, monthSpanish == null || monthSpanish.isBlank());
            if (sheet == null) {
                return result;
            }

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            String currentLabel = null;
            Double lastCost = null;
            CellStyle lastStyle = null;

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell b = row.getCell(1);
                if (b != null && b.getCellStyle().getFillForegroundColorColor() != null) {
                    if (currentLabel != null) {
                        result.add(new ServiceTeamRaw(currentLabel, lastCost, lastStyle));
                    }
                    currentLabel = b.getStringCellValue().trim();
                    lastCost = null;
                    lastStyle = null;
                }

                Cell h = row.getCell(7);
                if (h != null && currentLabel != null) {
                    CellValue val = evaluator.evaluate(h);
                    lastStyle = h.getCellStyle();
                    if (val != null && val.getCellType() == CellType.NUMERIC) {
                        lastCost = val.getNumberValue();
                    }
                }

                boolean emptyRow = true;
                for (int c = 0; c < 8; c++) {
                    Cell cc = row.getCell(c);
                    if (cc != null && cc.getCellType() != CellType.BLANK && !cc.toString().trim().isEmpty()) {
                        emptyRow = false;
                        break;
                    }
                }

                if (emptyRow && currentLabel != null) {
                    result.add(new ServiceTeamRaw(currentLabel, lastCost, lastStyle));
                    currentLabel = null;
                    lastCost = null;
                    lastStyle = null;
                }
            }

            if (currentLabel != null) {
                result.add(new ServiceTeamRaw(currentLabel, lastCost, lastStyle));
            }
        }

        result.removeIf(item -> item.getLabel() == null || item.getLabel().trim().isEmpty());
        return result;
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