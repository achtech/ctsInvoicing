package global;

import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

    public static class ServiceTeamRaw {
        private String label;
        private Double cost;
        private CellStyle style;

        public ServiceTeamRaw(String label, Double cost,CellStyle style) {
            this.label = label;
            this.cost = cost;
            this.style = style;
        }

        public String getLabel() { return label; }
        public Double getCost() { return cost; }
        public  CellStyle getStyle() {return style;}
    }

    public List<String> extractColumnBBlueValues(File file) throws Exception {
        List<String> result = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            for (Row row : sheet) {
                Cell cell = row.getCell(1); // Column B

                if (cell == null) continue;

                String text = cell.getStringCellValue().trim();
                if (text.isEmpty()) continue;

                CellStyle style = cell.getCellStyle();
                Color bgColor = style.getFillForegroundColorColor();

                if (bgColor != null) {
                    result.add(text);
                }
            }
        }
        return result;
    }

    public List<ServiceTeamRaw> extractRawServiceTeams(File file) throws Exception {
        List<ServiceTeamRaw> result = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
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
                        result.add(new ServiceTeamRaw(currentLabel, lastCost,lastStyle));
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
        return result;
    }
}
