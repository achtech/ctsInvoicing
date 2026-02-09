package forecast.by.month;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;

public class VacationTracker {

    private static final int START_ROW = 9;     // Excel row 10
    private static final int HEADER_ROW = 10;    // Excel row 11
    private static final int DATA_START_ROW = 11;
    private static final int TEAM_SIZE = 30;
    private static final int INITIAL_BALANCE = 25;

    public static void main(String[] args) throws Exception {

        Workbook wb = new XSSFWorkbook();

        createSheet(wb, "FY25", LocalDate.of(2025, 2, 1), 2);
        createSheet(wb, "FY26", LocalDate.of(2025, 4, 1), 12);

        try (FileOutputStream fos = new FileOutputStream("Vacation_Tracker_FY25_FY26.xlsx")) {
            wb.write(fos);
        }
        wb.close();
    }

    private static void createSheet(Workbook wb, String name, LocalDate startDate, int months) {
        Sheet sheet = wb.createSheet(name);

        // Styles
        CellStyle centerBold = wb.createCellStyle();
        centerBold.setAlignment(HorizontalAlignment.CENTER);
        centerBold.setVerticalAlignment(VerticalAlignment.CENTER);
        Font bold = wb.createFont();
        bold.setBold(true);
        centerBold.setFont(bold);

        CellStyle weekendStyle = wb.createCellStyle();
        weekendStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        weekendStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Legend
        createLegend(sheet, wb);

        // Fixed headers
        Row header = sheet.createRow(HEADER_ROW);
        header.createCell(0).setCellValue("Current project");
        header.createCell(1).setCellValue("Team Member");
        header.createCell(2).setCellValue("Balance for current FY");

        for (int i = 0; i <= 2; i++) {
            header.getCell(i).setCellStyle(centerBold);
            sheet.setColumnWidth(i, 5000);
        }

        // Team members
        for (int i = 0; i < TEAM_SIZE; i++) {
            Row row = sheet.createRow(DATA_START_ROW + i);
            row.createCell(1).setCellValue("Team Member " + (i + 1));
        }

        int col = 3;
        LocalDate dateCursor = startDate;

        for (int m = 0; m < months; m++) {
            YearMonth ym = YearMonth.from(dateCursor);
            int days = ym.lengthOfMonth();
            int monthStartCol = col;

            // Month header
            Row monthRow = sheet.getRow(START_ROW);
            if (monthRow == null) monthRow = sheet.createRow(START_ROW);

            Cell monthCell = monthRow.createCell(monthStartCol);
            monthCell.setCellValue(ym.getMonth() + " " + ym.getYear());
            monthCell.setCellStyle(centerBold);

            sheet.addMergedRegion(new CellRangeAddress(
                    START_ROW, START_ROW, monthStartCol, monthStartCol + days - 1));

            // Day numbers + weekends
            for (int d = 1; d <= days; d++) {
                Row dayRow = sheet.getRow(HEADER_ROW);
                Cell dayCell = dayRow.createCell(col);
                dayCell.setCellValue(d);
                dayCell.setCellStyle(centerBold);

                LocalDate current = ym.atDay(d);
                boolean weekend = current.getDayOfWeek() == DayOfWeek.SATURDAY
                        || current.getDayOfWeek() == DayOfWeek.SUNDAY;

                if (weekend) {
                    for (int r = HEADER_ROW; r < DATA_START_ROW + TEAM_SIZE; r++) {
                        Row wr = sheet.getRow(r);
                        if (wr == null) wr = sheet.createRow(r);
                        Cell wc = wr.createCell(col);
                        wc.setCellStyle(weekendStyle);
                    }
                }
                col++;
            }
            dateCursor = dateCursor.plusMonths(1);
        }

        // Conditional formatting
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

        addConditionalRule(scf, sheet, "V", IndexedColors.LIGHT_GREEN);
        addConditionalRule(scf, sheet, "S", IndexedColors.RED);
        addConditionalRule(scf, sheet, "L", IndexedColors.YELLOW);

        // Balance formulas
        String lastCol = CellReference.convertNumToColString(col - 1);

        for (int i = 0; i < TEAM_SIZE; i++) {
            int excelRow = DATA_START_ROW + i + 1; // Excel is 1-based

            sheet.getRow(DATA_START_ROW + i)
                    .createCell(2)
                    .setCellFormula(
                            INITIAL_BALANCE +
                                    "-COUNTIF(D" + excelRow + ":" + lastCol + excelRow + ",\"V\")"
                    );
        }
    }

    private static void addConditionalRule(SheetConditionalFormatting scf, Sheet sheet,
                                           String value, IndexedColors color) {

        ConditionalFormattingRule rule =
                scf.createConditionalFormattingRule(
                        ComparisonOperator.EQUAL, "\"" + value + "\"");

        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(color.getIndex());
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] range = {
                CellRangeAddress.valueOf("D12:ZZ100")
        };

        scf.addConditionalFormatting(range, rule);
    }

    private static void createLegend(Sheet sheet, Workbook wb) {
        Row r2 = sheet.createRow(1);
        r2.createCell(0).setCellValue("Legend:");

        Object[][] data = {
                {"V", "Vacation", IndexedColors.LIGHT_GREEN},
                {"S", "Sick Leave", IndexedColors.RED},
                {"L", "Legal Absence", IndexedColors.YELLOW},
                {"", "Weekend", IndexedColors.GREY_25_PERCENT}
        };

        for (int i = 0; i < data.length; i++) {
            Row r = sheet.createRow(2 + i);
            r.createCell(0).setCellValue((String) data[i][0]);
            r.createCell(1).setCellValue((String) data[i][1]);

            CellStyle cs = wb.createCellStyle();
            cs.setFillForegroundColor(((IndexedColors) data[i][2]).getIndex());
            cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            r.getCell(0).setCellStyle(cs);
        }
    }
}

