package forecast.by.month;

import forecast.by.month.service.ExcelFileNameGenerator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.time.Month;
import java.time.YearMonth;
import java.util.Locale;

public class Helper {

    // Helper method to check if a row is empty
    public static boolean isRowEmpty(Row row) {
        if (row == null)
            return true;
        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;

    }

    public static void writeWorkbook(Workbook workbook, String fileName) throws IOException {
        File directory = new File(new File(fileName).getParent());
        if (!directory.exists()) {
            if (!directory.mkdirs()) {
                throw new IOException("Failed to create directory: " + directory.getAbsolutePath());
            }
        }

        int retries = 5;
        long delayMs = 2000;
        for (int attempt = 1; attempt <= retries; attempt++) {
            try (FileOutputStream fos = new FileOutputStream(fileName)) {
                workbook.write(fos);
                System.out.println("Created file: " + fileName);
                return;
            } catch (FileNotFoundException e) {
                if (e.getMessage().contains("being used by another process") && attempt < retries) {
                    System.err.println("File locked: " + fileName + ". Retrying (" + attempt + "/" + retries
                            + ") after " + delayMs + "ms...");
                    try {
                        Thread.sleep(delayMs);
                    } catch (InterruptedException ie) {
                        Thread.currentThread().interrupt();
                        throw new IOException("Interrupted while waiting to retry writing file: " + fileName, ie);
                    }
                } else {
                    String tempFileName = fileName.replace(".xlsx", "_" + System.currentTimeMillis() + ".xlsx");
                    System.err.println(
                            "All retries failed for: " + fileName + ". Writing to temporary file: " + tempFileName);
                    try (FileOutputStream fos = new FileOutputStream(tempFileName)) {
                        workbook.write(fos);
//                        System.out.println("Created temporary file: " + tempFileName);
                        return;
                    }
                }
            }
        }
        throw new IOException("Failed to write file after " + retries + " attempts: " + fileName);
    }

    public static String getDesktopPath() {
        // Get the user's home directory
        String userHome = System.getProperty("user.home");
        if (userHome == null) {
            return null;
        }

        // Append the platform-specific desktop folder name
        String desktopFolder = "Desktop";
        // On some Linux systems, the desktop folder might be localized (e.g., "Escritorio" in Spanish)
        // For simplicity, assume "Desktop" (common on Windows, Mac, and most Linux distributions)

        return Path.of(userHome, desktopFolder).toString();
    }

    public static CellStyle getCenterStandardStyle(Workbook outputWorkbook) {
        // Create currency style for cost column
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getRightStandardStyle(Workbook outputWorkbook) {
        // Create currency style for cost column
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.RIGHT);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getLeftStandardStyle(Workbook outputWorkbook) {
        // Create currency style for cost column
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.LEFT);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getCurrencyStyle(Workbook outputWorkbook) {
        // Create currency style for cost column
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        DataFormat dataFormat = outputWorkbook.createDataFormat();
        headerStyle.setDataFormat(dataFormat.getFormat("#,##0.00 €"));
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.RIGHT);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getWeekendStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 128, (byte) 128}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getHeaderStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0, (byte) 51, (byte) 153}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        return headerStyle;
    }

    public static CellStyle getLegalAbsenceStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 0, (byte) 0}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getSickLeaveStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 0, (byte) 0}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getFreedayStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0, (byte) 0, (byte) 255}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getVacanceStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0, (byte) 255, (byte) 0}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getDateStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 128, (byte) 128, (byte) 128}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static CellStyle getFooterCurrencyStyle(Workbook outputWorkbook) {
        CellStyle headerStyle = outputWorkbook.createCellStyle();
        Font headerFont = outputWorkbook.createFont();
        DataFormat dataFormat = outputWorkbook.createDataFormat();
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0, (byte) 51, (byte) 153}, null));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setDataFormat(dataFormat.getFormat("#,##0.00 €"));
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.RIGHT);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        return headerStyle;
    }

    public static int numberOfDays(String sheetName) {
        String monthName = sheetName.substring(ExcelFileNameGenerator.SHEET_SERVICE_HOURS_DETAILS.length()).trim();
        try {
            int year = java.time.LocalDate.now().getYear();
            Locale locale = Locale.ENGLISH;
            int month = java.time.Month.from(
                    YearMonth.parse(
                            year + "-" + monthName.substring(0, 1).toUpperCase() + monthName.substring(1).toLowerCase(),
                            java.time.format.DateTimeFormatter.ofPattern("yyyy-MMMM")
                    ).atDay(1)
            ).getValue();
            return YearMonth.of(year, month).lengthOfMonth();
        } catch (Exception e) {
            System.err.println("Invalid month name: " + monthName);
            return -1; // Or throw an exception
        }
    }

    public static Double getRates(String input) {
        String[] parts = input.split(" > ");
        String[] lastSectionParts = parts[parts.length - 1].split(" - ", 2);
        String[] currencyParts = lastSectionParts[1].split("/");
        String oldRate = currencyParts.length == 1 ?currencyParts[0] : currencyParts[1];
        String result = oldRate.replace(" EUR", "").replace(",", ".").trim();
        result = result.contains(" (Operative)") ? result.replace(" (Operative)","") : result.contains(" (Ajuste)") ? result.replace(" (Ajuste)","") : result;
        result = result.contains(" - Operativa") ? result.replace(" - Operativa","") : result.contains(" (Ajuste)") ? result.replace(" (Ajuste)","") : result;
        result = result.contains(" - Operative") ? result.replace(" - Operative","") : result.contains(" (Ajuste)") ? result.replace(" (Ajuste)","") : result;
        result = result.contains("Tarifa ") ? result.replace("Tarifa ","") : result;
        try {
            return Double.parseDouble(result);
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("Failed to parse EUR value: " + result, e);
        }
    }

    public static String getColumnLetter(int columnIndex) {
        StringBuilder columnLetter = new StringBuilder();
        while (columnIndex >= 0) {
            columnLetter.insert(0, (char) ('A' + (columnIndex % 26)));
            columnIndex = (columnIndex / 26) - 1;
        }
        return columnLetter.toString();
    }

    public static int getMonthFromSheetName(String invoicingSheetName) {
        String monthName = invoicingSheetName.substring(22);
        Month month = Month.valueOf(monthName.toUpperCase());
        // Get the month number (1-12)
        return month.getValue();
    }


}
