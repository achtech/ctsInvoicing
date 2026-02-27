package invoicing.Helper;
// ReferenceData.java
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReferenceData {
    private Map<String, BigDecimal> groupToRate = new HashMap<>();
    private Map<BigDecimal, String> rateToGroup = new HashMap<>();

    public void load(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path)) {
            load(fis);
        }
    }

    public void load(InputStream is) throws IOException {
        try (Workbook wb = new XSSFWorkbook(is)) {
            Sheet sheet = wb.getSheetAt(0); // Reference Table: Column A = GroupId, Column B = Rate
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header
                String groupId = getStringCellValue(row.getCell(0)); // Column A
                BigDecimal rate = getBigDecimalFromCell(row.getCell(1));    // Column B
                if (groupId != null && rate!=null && rate.compareTo(BigDecimal.ZERO) > 0) {
                    groupToRate.put(groupId, rate);
                    rateToGroup.put(rate, groupId);
                }
            }
        }
    }

    private static BigDecimal getBigDecimalFromCell(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return null;
        }

        // If the cell contains a formula, we evaluate the result type
        CellType type = (cell.getCellType() == CellType.FORMULA)
                ? cell.getCachedFormulaResultType()
                : cell.getCellType();

        if (type == CellType.NUMERIC) {
            // Convert double to String first to maintain precision in BigDecimal
            return new BigDecimal(String.valueOf(cell.getNumericCellValue()));
        } else if (type == CellType.STRING) {
            try {
                return new BigDecimal(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                // Handle cases where the string isn't a valid number
                return null;
            }
        }

        return null;
    }


    public BigDecimal getRateByGroup(String group) {
        return groupToRate.get(group);
    }

    /**
     * Finds a GroupId by matching the rate approximately (within a small tolerance).
     * Returns null if no match is found.
     */
    public String getGroupByApproximateRate(BigDecimal rate) {
        double tolerance = 0.5; // Allow small differences
        double rateFromInvoice = rate.doubleValue();
        // Try exact match first
        if (rateToGroup.containsKey(rateFromInvoice)) {
            return rateToGroup.get(rateFromInvoice);
        }
        
        // Find closest match within tolerance
        String bestMatch = null;
        double smallestDiff = tolerance;
        
        for (Map.Entry<String, BigDecimal> entry : groupToRate.entrySet()) {
            double diff = Math.abs(entry.getValue().doubleValue() - rateFromInvoice);
            if (diff < smallestDiff) {
                smallestDiff = diff;
                bestMatch = entry.getKey();
            }
        }
        
        return bestMatch;
    }
    
    public BigDecimal getCorrectRateByApproximate(BigDecimal rate) {
        String groupId = getGroupByApproximateRate(rate);
        return groupId != null ? groupToRate.get(groupId) : null;
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue());
        }
        return null;
    }

    private static double getDoubleCellValue(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        return 0.0;
    }
}
