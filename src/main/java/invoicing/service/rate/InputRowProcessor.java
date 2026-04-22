package invoicing.service.rate;

// InputRowProcessor.java
import invoicing.Helper.ReferenceData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class InputRowProcessor {

    private final ReferenceData referenceData;
    // Store processed data for debugging
    private final List<ProcessedDataEntry> processedData = new ArrayList<>();
    static int count = 0;

    // Container for processed data
    private static class ProcessedDataEntry {
        String groupId;
        BigDecimal rateFromInvoice;
        BigDecimal correctRate;
        double hours;
        BigDecimal invoicedRate;
        BigDecimal cost;

        ProcessedDataEntry(String groupId, BigDecimal rateFromInvoice, BigDecimal correctRate,
                          double hours, BigDecimal invoicedRate, BigDecimal cost) {
            this.groupId = groupId;
            this.rateFromInvoice = rateFromInvoice;
            this.correctRate = correctRate;
            this.hours = hours;
            this.invoicedRate = invoicedRate;
            this.cost = cost;
        }
    }

    public InputRowProcessor(ReferenceData referenceData) {
        this.referenceData = referenceData;
    }

    /**
     * Prints the processed data for debugging.
     */
    public void printProcessedData() {
        System.out.println("\n=== PROCESSED DATA (Ordered by GroupId) ===");
        
        processedData.sort(Comparator.comparing(e -> e.groupId));
        
        for (ProcessedDataEntry data : processedData) {
            System.out.println("GroupId: " + (data.groupId != null ? data.groupId : "Unknown") + 
                               " | RateFromInvoice: " + String.format("%.2f", data.rateFromInvoice) +
                               " | CorrectRate: " + String.format("%.2f", data.correctRate) +
                               " | Hours: " + String.format("%.2f", data.hours) + 
                               " | InvoicedRate: " + String.format("%.2f", data.invoicedRate) +
                               " | Cost: " + String.format("%.2f", data.cost));
        }
        System.out.println("=============================================\n");
    }

    /**
     * Processes one row from input file - simplified version.
     * Uses raw cost from column H and only uses rate for GroupId mapping.
     */
    public RowData processRow(Row row, FormulaEvaluator evaluator) {
        if (row == null) return null;

        // Extract rate from column C text (for GroupId mapping)
        String columnCText = getStringCellValue(row.getCell(2)); // Column C
        BigDecimal columnG = getBigDecimalFromCell(row.getCell(6), evaluator); // Column G

        BigDecimal rate = extractRateFromText(columnCText);
        
        // Fallback to column G if no rate in text
        if (rate == null || rate.compareTo(BigDecimal.ZERO) == 0) {
            if (columnG != null && columnG.compareTo(BigDecimal.ZERO) != 0) {
                rate = columnG;
            }
        }

        // Handle special rates (Guardias)
        String groupId = null;
        BigDecimal referenceRate = null;
        
        // Get cost from column H first
        BigDecimal cost = getBigDecimalFromCell(row.getCell(7), evaluator);
        if (cost == null) {
            cost = BigDecimal.ZERO;
        }
        
        if (rate != null && rate.compareTo(BigDecimal.ZERO) != 0) {
            if (rate.doubleValue() == 25.0) {
                groupId = "Guardias - 25";
                referenceRate = new BigDecimal("25.00");
            } else if (rate.doubleValue() == 50.0) {
                groupId = "Guardias-50";
                referenceRate = new BigDecimal("50.00");
            } else {
                // Look up GroupId from reference data
                groupId = referenceData.getGroupByApproximateRate(rate);
                if (groupId != null) {
                    referenceRate = referenceData.getRateByGroup(groupId);
                } else {
                    // Rate not found in ReferenceData - create a placeholder group
                    groupId = "Other_" + rate.setScale(2, java.math.RoundingMode.HALF_UP).toPlainString().replace(".", "_");
                    referenceRate = rate;
                }
            }
} else if (cost.compareTo(BigDecimal.ZERO) != 0) {
            // Check if this is a total/summary row - column B is empty
            String colB = getStringCellValue(row.getCell(1)); // Column B - employee name or label
            if (colB == null || colB.trim().isEmpty()) {
                return null; // Skip summary/total rows
            }
            // Rate is 0 or null, but cost > 0 - these are Tools/Other costs (expenses)
            groupId = "Tools";
            referenceRate = BigDecimal.ONE;
        }
        
        // If still no groupId, skip this row
        if (groupId == null) {
            return null;
        }

        // Extract hours from column D
        double hours = getDoubleCellValue(row.getCell(3));

        // For Tools/expenses (rate=0), hours should be 0 - expenses are cost only
        if ("Tools".equals(groupId)) {
            hours = 0;
        } else if (hours == 0 && cost.doubleValue() != 0 && referenceRate != null && referenceRate.doubleValue() != 0) {
            // Only for regular groups calculate hours from cost
            hours = cost.doubleValue() / referenceRate.doubleValue();
        }

        // Store for debugging
        processedData.add(new ProcessedDataEntry(groupId, rate, referenceRate, 
                                                  hours, columnG, cost));

        return new RowData(groupId, hours, cost.doubleValue());
    }

    /**
     * Process row without evaluator (for backward compatibility)
     */
    public RowData processRow(Row row) {
        return processRow(row, null);
    }

    /**
     * Extracts the EUR rate from text like:
     * "Horas servicio: Tarifa 195,00 MAD/18,08 EUR (Operative)" => 18.08
     * "Horas servicio: Tarifa 25,43 EUR (Operative)" => 25.43
     */
    private BigDecimal extractRateFromText(String text) {
        if (text == null || text.trim().isEmpty()) {
            return new BigDecimal(0);
        }

        // Pattern to match number before "EUR"
        // Handles both comma and dot as decimal separators
        Pattern pattern = Pattern.compile("([0-9]+[,.]?[0-9]*)\\s*EUR");
        Matcher matcher = pattern.matcher(text);
        
        if (matcher.find()) {
            String rateStr = matcher.group(1).replace(',', '.');
            try {
                return new BigDecimal(rateStr);
            } catch (NumberFormatException e) {
                System.err.println("Warning: Could not parse rate from: " + text);
                return new BigDecimal(0);
            }
        }
        return new BigDecimal(0);
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        }
        return null;
    }

    private static double getDoubleCellValue(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            String str = cell.getStringCellValue().trim().replace(',', '.');
            try {
                return Double.parseDouble(str);
            } catch (NumberFormatException e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private static BigDecimal getBigDecimalFromCell(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return null;
        }

        CellType cellType = cell.getCellType();
        
        // If it's a formula, evaluate it
        if (cellType == CellType.FORMULA && evaluator != null) {
            CellValue evaluatedValue = evaluator.evaluate(cell);
            if (evaluatedValue != null && evaluatedValue.getCellType() == CellType.NUMERIC) {
                return new BigDecimal(String.valueOf(evaluatedValue.getNumberValue()));
            }
            return null;
        }

        if (cellType == CellType.BLANK) {
            return null;
        }

        if (cellType == CellType.NUMERIC) {
            return new BigDecimal(String.valueOf(cell.getNumericCellValue()));
        } else if (cellType == CellType.STRING) {
            try {
                String str = cell.getStringCellValue().trim();
                str = str.replace("EUR", "").replace("MAD", "").trim();
                str = str.replace(",", ".");
                return new BigDecimal(str);
            } catch (NumberFormatException e) {
                return null;
            }
        }

        return null;
    }

    private static BigDecimal getBigDecimalFromCell(Cell cell) {
        return getBigDecimalFromCell(cell, null);
    }
    private static boolean isNotBlank(String str) {
        return str != null && !str.trim().isEmpty();
    }

    /**
     * Data class for returning processed row data.
     */
    public static class RowData {
        private final String groupId;
        private final double hours;
        private final double cost;

        public RowData(String groupId, double hours, double cost) {
            this.groupId = groupId;
            this.hours = hours;
            this.cost = cost;
        }

        public String getGroupId() {
            return groupId;
        }

        public double getHours() {
            return hours;
        }

        public double getCost() {
            return cost;
        }
    }
}
