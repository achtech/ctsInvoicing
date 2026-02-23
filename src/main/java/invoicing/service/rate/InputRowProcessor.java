package invoicing.service.rate;

// InputRowProcessor.java
import invoicing.Helper.ReferenceData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
     * Processes one row from input file according to the new workflow.
     * Returns RowData with GroupId, hours, and cost.
     */
    public RowData processRow(Row row) {
        if (row == null) return null;

        // Step 3.1: Extract rate from column C (parse text)
        String columnBText = getStringCellValue(row.getCell(1)); // Column C (index 2)
        String columnCText = getStringCellValue(row.getCell(2)); // Column C (index 2)
        double columnDText = getDoubleCellValue(row.getCell(3)); // Column C (index 2)
        BigDecimal columnGText = getBigDecimalFromCell(row.getCell(6)); // Column C (index 2)
        BigDecimal columnHText = getBigDecimalFromCell(row.getCell(7)); // Column C (index 2)

        BigDecimal rateFromInvoice = extractRateFromText(columnCText);
        if(BigDecimal.ZERO.compareTo(rateFromInvoice) == 0){
            if(columnGText!=null && BigDecimal.ZERO.compareTo(columnGText) != 0){
                rateFromInvoice = columnGText;
            }
            else {
                if(columnBText!=null && columnHText!=null && !columnBText.isEmpty() && !columnCText.isEmpty() && columnHText.compareTo(BigDecimal.ZERO) !=0 ){
                    rateFromInvoice = new BigDecimal("1");
                }
            }
        }

        // Step 3.2: Look up correct rate and GroupId from reference data
        String groupId = referenceData.getGroupByApproximateRate(rateFromInvoice);
        BigDecimal correctRate = referenceData.getCorrectRateByApproximate(rateFromInvoice);
        if (groupId == null || correctRate == null) {
            System.out.println("Warning: No matching GroupId found for rate: " + rateFromInvoice);
            return null;
        }

        // Step 3.3: Extract hours from column D
        double hours = getDoubleCellValue(row.getCell(3)); // Column D (index 3)

        // Step 3.4: Extract invoicedRate from column G
        BigDecimal invoicedRate = getBigDecimalFromCell(row.getCell(6)); // Column G (index 6)

        // Step 3.5: Extract Cost from column H
        BigDecimal cost = getBigDecimalFromCell(row.getCell(7)); // Column H (index 7)
        if(hours == 0 && correctRate != null && correctRate.doubleValue() == 1){
            hours =correctRate.doubleValue();
        }
        // Step 3.6: Apply correction logic
        if (Math.abs(invoicedRate.doubleValue() - correctRate.doubleValue()) > 0.01) { // invoicedRate != correctRate
            if (invoicedRate.compareTo(BigDecimal.ZERO) == 0.0 && cost.compareTo(BigDecimal.ZERO) != 0.0) {
                // Case a: invoicedRate = 0 and Cost != 0
                hours = cost.doubleValue();
                invoicedRate = new BigDecimal(1.0);
                cost = new BigDecimal(hours * invoicedRate.doubleValue()); // Recalculate cost
            } else if (invoicedRate == new BigDecimal(25.0) || invoicedRate == new BigDecimal(50.0)) {
                // Case c: Special Guardas cases
                correctRate = invoicedRate;
                groupId = invoicedRate == new BigDecimal(25.0) ? "Guardas_25" : "Guardas_50";
                cost = new BigDecimal(hours * correctRate.doubleValue()); // Recalculate cost
            } else {
                // Case b: General case - adjust hours and use correct rate
                hours = hours * invoicedRate.doubleValue() / correctRate.doubleValue();
                invoicedRate = correctRate;
                cost = new BigDecimal(hours * invoicedRate.doubleValue());
            }
        } else {
            // invoicedRate == correctRate, recalculate cost to ensure consistency
            cost = new BigDecimal(hours * correctRate.doubleValue());
        }

        // Store for debugging
        processedData.add(new ProcessedDataEntry(groupId, rateFromInvoice, correctRate, 
                                                  hours, invoicedRate, cost));

        return new RowData(groupId, hours, cost.doubleValue());
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
