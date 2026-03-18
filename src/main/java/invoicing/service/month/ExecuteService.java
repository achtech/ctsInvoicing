package invoicing.service.month;

import invoicing.Helper.Helper;
import invoicing.service.month.impl.*;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.ArrayList;
import java.util.List;

import static invoicing.service.month.ExcelFileNameGenerator.*;

public class ExecuteService {
    private final DateProvider dateProvider;
    private final ExcelFileNameGenerator fileNameGenerator;
    private final ExcelReader excelReader;
    private final ServiceTeamExtractor serviceTeamExtractor;
    private final ExcelWriter excelWriter;

    public ExecuteService(
            DateProvider dateProvider,
            ExcelFileNameGenerator fileNameGenerator,
            ExcelReader excelReader,
            ServiceTeamExtractor serviceTeamExtractor,
            ExcelWriter excelWriter) {
        this.dateProvider = dateProvider;
        this.fileNameGenerator = fileNameGenerator;
        this.excelReader = excelReader;
        this.serviceTeamExtractor = serviceTeamExtractor;
        this.excelWriter = excelWriter;
    }

    public static final class SheetTableData {
        private final String sheetName;
        private final List<List<String>> rows;

        public SheetTableData(String sheetName, List<List<String>> rows) {
            this.sheetName = sheetName;
            this.rows = rows;
        }

        public String getSheetName() {
            return sheetName;
        }

        public List<List<String>> getRows() {
            return rows;
        }

        public int getRowCount() {
            return rows == null ? 0 : rows.size();
        }

        public int getColumnCount() {
            if (rows == null) return 0;
            int max = 0;
            for (List<String> r : rows) {
                if (r != null && r.size() > max) max = r.size();
            }
            return max;
        }

        @Override
        public String toString() {
            return "SheetTableData{" +
                    "sheetName='" + sheetName + '\'' +
                    ", rows=" + getRowCount() +
                    ", cols=" + getColumnCount() +
                    '}';
        }
    }

    public static final class MonthWorkbookData {
        private final String serviceTeam;
        private final List<SheetTableData> sheets;

        public MonthWorkbookData(String serviceTeam, List<SheetTableData> sheets) {
            this.serviceTeam = serviceTeam;
            this.sheets = sheets;
        }

        public String getServiceTeam() {
            return serviceTeam;
        }

        public List<SheetTableData> getSheets() {
            return sheets;
        }

        public int getSheetCount() {
            return sheets == null ? 0 : sheets.size();
        }

        @Override
        public String toString() {
            return "MonthWorkbookData{" +
                    "serviceTeam='" + serviceTeam + '\'' +
                    ", sheets=" + getSheetCount() +
                    '}';
        }
    }

    public void process(String inputExcelFilePath, String outputGeneratedExcelsFilePath, int monthsToProcess) {
        // Normalize the path for cross-platform compatibility
        File file = new File(inputExcelFilePath).getAbsoluteFile();

        // Validate file
        if (!file.exists()) {
            System.err.println("Error: File does not exist: " + file.getAbsolutePath());
            return;
        }
        if (!file.isFile()) {
            System.err.println("Error: Path is not a file: " + file.getAbsolutePath());
            return;
        }
        if (!file.getName().toLowerCase().endsWith(".xlsx") && !file.getName().toLowerCase().endsWith(".xls")) {
            System.err.println("Error: File is not an Excel file (.xlsx or .xls): " + file.getAbsolutePath());
            return;
        }

        try (Workbook inputWorkbook = WorkbookFactory.create(file)) {

            LocalDate currentDate = dateProvider.getCurrentDate(inputExcelFilePath);
            int currentYear = dateProvider.getYear(currentDate);
            int currentMonth = dateProvider.getMonthValue(currentDate);
            String currentMonthSpanish = dateProvider.getMonthNameSpanish(currentDate);

            // Collect all projects/service teams across the requested months.
            // Some projects can start next month, so they may not appear in the current Facturación sheet.
            java.util.Set<String> serviceTeamNamesSet = new java.util.LinkedHashSet<>();
            for (int i = 0; i < monthsToProcess; i++) {
                LocalDate dateForSheet = currentDate.plusMonths(i);
                String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);
                String facturacionName = SHEET_FACTURACIÓN + " " + monthNameSpa;
                Sheet facturacionSheet = excelReader.getSheet(inputWorkbook, facturacionName);
                if (facturacionSheet == null) {
                    System.err.println("Warning: Sheet not found: " + facturacionName);
                    continue;
                }
                List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
                serviceTeamNamesSet.addAll(serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames));
            }

            List<String> serviceTeamNames = new java.util.ArrayList<>(serviceTeamNamesSet);

            // 1. GENERATE SEPARATE FILES
            for (String serviceTeam : serviceTeamNames) {
                // Collect month names for this run
                List<String> monthNames = new ArrayList<>();
                for (int i = 0; i < monthsToProcess; i++) {
                    monthNames.add(dateProvider.getMonthNameEnglish(currentDate.plusMonths(i)));
                }

                try (Workbook outputWorkbook = excelWriter.createWorkbookWithSheets(monthNames)) {
                    for (int i = 0; i < monthsToProcess; i++) {
                        LocalDate dateForSheet = currentDate.plusMonths(i);
                        String monthNameEng = dateProvider.getMonthNameEnglish(dateForSheet);
                        String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);

                        excelWriter.copyServiceHoursSheetData(
                                inputWorkbook,
                                outputWorkbook,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + monthNameSpa,
                                SHEET_SERVICE_HOURS_DETAILS + " " + monthNameEng,
                                SHEET_AJUSTES,
                                SHEET_FACTURACIÓN + " " + monthNameSpa
                        );
                    }

                    String outputFileName = fileNameGenerator.generateOutputFileName(currentMonth, currentYear, serviceTeam, outputGeneratedExcelsFilePath);
                    Helper.writeWorkbook(outputWorkbook, outputFileName);
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public List<MonthWorkbookData> processToData(String inputExcelFilePath, int monthsToProcess) throws IOException {
        File file = new File(inputExcelFilePath).getAbsoluteFile();
        if (!file.exists() || !file.isFile()) {
            throw new IllegalArgumentException("Invalid Excel file: " + file.getAbsolutePath());
        }

        try (Workbook inputWorkbook = WorkbookFactory.create(file)) {
            LocalDate currentDate = dateProvider.getCurrentDate(inputExcelFilePath);

            java.util.Set<String> serviceTeamNamesSet = new java.util.LinkedHashSet<>();
            for (int i = 0; i < monthsToProcess; i++) {
                LocalDate dateForSheet = currentDate.plusMonths(i);
                String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);
                String facturacionName = SHEET_FACTURACIÓN + " " + monthNameSpa;
                Sheet facturacionSheet = excelReader.getSheet(inputWorkbook, facturacionName);
                if (facturacionSheet == null) continue;
                List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
                serviceTeamNamesSet.addAll(serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames));
            }

            if (serviceTeamNamesSet.isEmpty()) return Collections.emptyList();

            List<String> serviceTeamNames = new ArrayList<>(serviceTeamNamesSet);
            List<String> monthNames = new ArrayList<>();
            for (int i = 0; i < monthsToProcess; i++) {
                monthNames.add(dateProvider.getMonthNameEnglish(currentDate.plusMonths(i)));
            }

            List<MonthWorkbookData> result = new ArrayList<>();

            for (String serviceTeam : serviceTeamNames) {
                try (Workbook outputWorkbook = excelWriter.createWorkbookWithSheets(monthNames)) {
                    for (int i = 0; i < monthsToProcess; i++) {
                        LocalDate dateForSheet = currentDate.plusMonths(i);
                        String monthNameEng = dateProvider.getMonthNameEnglish(dateForSheet);
                        String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);

                        excelWriter.copyServiceHoursSheetData(
                                inputWorkbook,
                                outputWorkbook,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + monthNameSpa,
                                SHEET_SERVICE_HOURS_DETAILS + " " + monthNameEng,
                                SHEET_AJUSTES,
                                SHEET_FACTURACIÓN + " " + monthNameSpa
                        );
                    }

                    result.add(new MonthWorkbookData(serviceTeam, workbookToTables(outputWorkbook)));
                }
            }

            return result;
        }
    }

    private List<SheetTableData> workbookToTables(Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        DataFormatter formatter = new DataFormatter();

        List<SheetTableData> sheets = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            List<List<String>> rows = new ArrayList<>();

            int lastRow = sheet.getLastRowNum();
            int maxCells = 0;
            for (int r = 0; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row != null && row.getLastCellNum() > maxCells) {
                    maxCells = row.getLastCellNum();
                }
            }

            for (int r = 0; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                List<String> cols = new ArrayList<>(maxCells);
                for (int c = 0; c < maxCells; c++) {
                    Cell cell = row == null ? null : row.getCell(c);
                    cols.add(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
                }
                rows.add(cols);
            }

            sheets.add(new SheetTableData(sheet.getSheetName(), rows));
        }
        return sheets;
    }

    public static void executeScript(String inputExcelFilePath, String outputExcelsFilePath, int monthsToProcess) throws Exception {
        ExecuteService processor = new ExecuteService(
                new DefaultDateProvider(),
                new DefaultExcelFileNameGenerator(),
                new DefaultExcelReader(),
                new ServiceTeamExtractorImpl(),
                new DefaultExcelWriter()
        );
        processor.process(inputExcelFilePath, outputExcelsFilePath, monthsToProcess);
    }

    public static List<MonthWorkbookData> executeToData(String inputExcelFilePath, int monthsToProcess) throws Exception {
        ExecuteService processor = new ExecuteService(
                new DefaultDateProvider(),
                new DefaultExcelFileNameGenerator(),
                new DefaultExcelReader(),
                new ServiceTeamExtractorImpl(),
                new DefaultExcelWriter()
        );
        return processor.processToData(inputExcelFilePath, monthsToProcess);
    }
}
