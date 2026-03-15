package invoicing.service.rate;

// InputRowProcessor.java
import invoicing.Helper.ReferenceData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.math.BigDecimal;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class InputRowProcessor {

    private final ReferenceData referenceData;

    public InputRowProcessor(ReferenceData referenceData) {
        this.referenceData = referenceData;
    }

    /**
     * Processes one row from input file according to the new workflow.
     * Returns RowData with GroupId, hours, and cost.
     */
    public RowData processRow(Row row) {
        if (row == null) return null;

        String columnBText = getStringCellValue(row.getCell(1));
        String columnCText = getStringCellValue(row.getCell(2));
        BigDecimal columnGText = getBigDecimalFromCell(row.getCell(6));
        BigDecimal columnHText = getBigDecimalFromCell(row.getCell(7));

        BigDecimal rateFromInvoice = extractRateFromText(columnCText);
        if (BigDecimal.ZERO.compareTo(rateFromInvoice) == 0) {
            if (columnGText != null && BigDecimal.ZERO.compareTo(columnGText) != 0) {
                rateFromInvoice = columnGText;
            } else if (columnBText != null && columnHText != null && !columnBText.isEmpty()
                    && columnCText != null && !columnCText.isEmpty()
                    && columnHText.compareTo(BigDecimal.ZERO) != 0) {
                rateFromInvoice = BigDecimal.ONE;
            }
        }

        String groupId = referenceData.getGroupByApproximateRate(rateFromInvoice);
        BigDecimal correctRate = referenceData.getCorrectRateByApproximate(rateFromInvoice);
        if (groupId == null || correctRate == null) {
            return null;
        }

        double hours = getDoubleCellValue(row.getCell(3));
        BigDecimal invoicedRate = getBigDecimalFromCell(row.getCell(6));
        BigDecimal cost = getBigDecimalFromCell(row.getCell(7));

        if (invoicedRate == null) invoicedRate = BigDecimal.ZERO;
        if (cost == null) cost = BigDecimal.ZERO;

        if (hours == 0 && correctRate.compareTo(BigDecimal.ONE) == 0) {
            hours = 1.0;
        }

        if (Math.abs(invoicedRate.doubleValue() - correctRate.doubleValue()) > 0.01) {
            if (invoicedRate.compareTo(BigDecimal.ZERO) == 0 && cost.compareTo(BigDecimal.ZERO) != 0) {
                // Case A: invoicedRate = 0 and Cost != 0
                hours = cost.doubleValue();
                invoicedRate = BigDecimal.ONE;
                cost = new BigDecimal(hours);
            } else if (invoicedRate.compareTo(new BigDecimal("25")) == 0
                    || invoicedRate.compareTo(new BigDecimal("50")) == 0) {
                // Case C: Special Guardas cases
                correctRate = invoicedRate;
                groupId = invoicedRate.compareTo(new BigDecimal("25")) == 0 ? "Guardas_25" : "Guardas_50";
                cost = new BigDecimal(hours * correctRate.doubleValue());
            } else {
                // Case B: General case
                hours = hours * invoicedRate.doubleValue() / correctRate.doubleValue();
                invoicedRate = correctRate;
                cost = new BigDecimal(hours * invoicedRate.doubleValue());
            }
        } else {
            // Rates match: recalc cost to ensure consistency
            cost = new BigDecimal(hours * correctRate.doubleValue());
        }

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
