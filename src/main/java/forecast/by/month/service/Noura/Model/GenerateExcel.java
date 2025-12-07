package forecast.by.month.service.Noura.Model;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class GenerateExcel {

    private static final String INPUT_EXCEL = "src/main/resources/Tickets.xlsx";
    private static final String TEMPLATE_EXCEL = "src/main/resources/Template.xlsx";
    private static final String OUTPUT_DIR = "src/main/resources/PR/";

    public static void main(String[] args) {
        try {
            List<Ticket> tickets = readTicketsFromExcel(INPUT_EXCEL);
            generateExcelFiles(tickets);
            System.out.println("✅ Excel files generated successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Step 1: Read the list of tickets from Tickets.xlsx
    public static List<Ticket> readTicketsFromExcel(String filePath) throws IOException {
        List<Ticket> tickets = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();

            // Skip header (assumed first row)
            for (int i = 1; i < rows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String ticketName = getCellValue(row.getCell(0));
                String prDate = getCellValue(row.getCell(1));
                String prResponsible = getCellValue(row.getCell(2));
                String prDeveloper = getCellValue(row.getCell(3));

                if (ticketName != null && !ticketName.isEmpty()) {
                    tickets.add(new Ticket(ticketName, prDate, prResponsible, prDeveloper));
                }
            }
        }
        return tickets;
    }

    // Step 2: Generate new Excel files based on Template.xlsx
    public static void generateExcelFiles(List<Ticket> tickets) throws IOException {
        for (Ticket ticket : tickets) {
            try (FileInputStream fis = new FileInputStream(TEMPLATE_EXCEL);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(1); // second sheet (index 1)

                // Fill required cells
                setCellValue(sheet, "D11", ticket.getTicketName());
                setCellValue(sheet, "D15", ticket.getPrDeveloper());
                setCellValue(sheet, "D17", ticket.getPrResponsible());
                setCellValue(sheet, "D19", ticket.getPrDate());

                // Save new file
                String outputFileName = OUTPUT_DIR + "Peer review " + ticket.getTicketName() + "_OK.xlsx";
                try (FileOutputStream fos = new FileOutputStream(outputFileName)) {
                    workbook.write(fos);
                }
            }
        }
    }

    // Helper: get cell value as string
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getDateCellValue().toString()
                    : String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    // Helper: convert "D11" to row/column indexes and set value
    private static void setCellValue(Sheet sheet, String cellRef, String value) {
        int col = cellRef.charAt(0) - 'A';
        int row = Integer.parseInt(cellRef.substring(1)) - 1;
        Row rowObj = sheet.getRow(row);
        if (rowObj == null) rowObj = sheet.createRow(row);
        Cell cell = rowObj.getCell(col);
        if (cell == null) cell = rowObj.createCell(col);
        cell.setCellValue(value);
    }
}

