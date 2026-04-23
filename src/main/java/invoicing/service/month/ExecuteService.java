package invoicing.service.month;

import invoicing.Helper.Helper;
import invoicing.service.month.impl.DefaultDateProvider;
import invoicing.service.month.impl.DefaultExcelFileNameGenerator;
import invoicing.service.month.impl.DefaultExcelReader;
import invoicing.service.month.impl.DefaultExcelWriter;
import invoicing.service.month.impl.ServiceTeamExtractorImpl;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.StringJoiner;
import java.util.function.Consumer;

public class ExecuteService {
    private static final String SHEET_AJUSTES = "Ajustes";
    private static final String SHEET_SERVICE_HOURS_DETAILS = "Service Hours Details";
    private static final String SHEET_HORAS_SERVICIO = "Horas servicio";
    private static final String SHEET_FACTURACION = "Facturaci\u00F3n";

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

    private static class MonthSpec {
        private final String monthNameEng;
        private final String monthNameSpa;
        private final String ajustesSheetName;

        private MonthSpec(String monthNameEng, String monthNameSpa, String ajustesSheetName) {
            this.monthNameEng = monthNameEng;
            this.monthNameSpa = monthNameSpa;
            this.ajustesSheetName = ajustesSheetName;
        }
    }

    public void process(String inputExcelFilePath, String outputGeneratedExcelsFilePath, int monthsToProcess) {
        process(inputExcelFilePath, outputGeneratedExcelsFilePath, monthsToProcess, null);
    }

    public void process(String inputExcelFilePath, String outputGeneratedExcelsFilePath, int monthsToProcess, Consumer<String> logger) {
        File file = new File(inputExcelFilePath).getAbsoluteFile();
        log(logger, "Resolved path: " + file.getAbsolutePath());

        if (!file.exists()) {
            log(logger, "Error: File does not exist: " + file.getAbsolutePath());
            return;
        }
        if (!file.isFile()) {
            log(logger, "Error: Path is not a file: " + file.getAbsolutePath());
            return;
        }
        if (!file.getName().toLowerCase().endsWith(".xlsx") && !file.getName().toLowerCase().endsWith(".xls")) {
            log(logger, "Error: File is not an Excel file (.xlsx or .xls): " + file.getAbsolutePath());
            return;
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook inputWorkbook = new XSSFWorkbook(fis)) {

            LocalDate currentDate = dateProvider.getCurrentDate(inputExcelFilePath);
            int currentYear = dateProvider.getYear(currentDate);
            String outputDirectory = outputGeneratedExcelsFilePath;

            Set<String> workbookMonths = new LinkedHashSet<>();
            String facturacionPrefix = SHEET_FACTURACION + " ";
            for (int s = 0; s < inputWorkbook.getNumberOfSheets(); s++) {
                String sheetName = inputWorkbook.getSheetName(s);
                if (!sheetName.startsWith(facturacionPrefix)) {
                    continue;
                }
                String monthSpa = sheetName.substring(facturacionPrefix.length()).trim().toLowerCase(Locale.ROOT);
                if (monthSpa.isEmpty()) {
                    continue;
                }
                String horasSheetName = SHEET_HORAS_SERVICIO + " " + monthSpa;
                if (inputWorkbook.getSheet(horasSheetName) != null) {
                    workbookMonths.add(monthSpa);
                }
            }

            List<String> workbookMonthList = new ArrayList<>(workbookMonths);
            if (workbookMonthList.isEmpty()) {
                log(logger, "No month sheet pairs found in file '" + file.getName() + "'. Expected '" + SHEET_FACTURACION + " <month>' and '" + SHEET_HORAS_SERVICIO + " <month>'.");
                return;
            }

            if (workbookMonthList.size() < monthsToProcess) {
                log(logger, "Requested " + monthsToProcess + " months for '" + file.getName() + "' but only found " + workbookMonthList.size() + ": " + String.join(", ", workbookMonthList));
            }

            int monthsToUse = Math.min(monthsToProcess, workbookMonthList.size());
            List<MonthSpec> availableMonths = new ArrayList<>();
            for (int i = 0; i < monthsToUse; i++) {
                String monthSpa = workbookMonthList.get(i);
                availableMonths.add(new MonthSpec(toEnglishMonthName(monthSpa), monthSpa, SHEET_AJUSTES));
            }

            if (availableMonths.isEmpty()) {
                log(logger, "No requested month sheets available in file '" + file.getName() + "'. Skipping file.");
                return;
            }

            String currentMonthSpanish = availableMonths.get(0).monthNameSpa;
            int currentMonth = getMonthNumber(currentMonthSpanish, currentDate.getMonthValue());

            Sheet facturacionSheet = inputWorkbook.getSheet(SHEET_FACTURACION + " " + availableMonths.get(0).monthNameSpa);
            if (facturacionSheet == null) {
                log(logger, "Unable to find a base Facturacion sheet in file '" + file.getName() + "'. Skipping file.");
                return;
            }

            List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
            List<String> serviceTeamNames = serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames);
            if (serviceTeamNames.isEmpty()) {
                log(logger, "No service teams found in file '" + file.getName() + "'.");
                return;
            }

            for (String serviceTeam : serviceTeamNames) {
                List<String> monthNames = new ArrayList<>();
                for (MonthSpec month : availableMonths) {
                    monthNames.add(month.monthNameEng);
                }

                try (Workbook outputWorkbook = excelWriter.createWorkbookWithSheets(monthNames)) {
                    for (MonthSpec month : availableMonths) {
                        excelWriter.copyServiceHoursSheetData(
                                inputWorkbook,
                                outputWorkbook,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + month.monthNameSpa,
                                SHEET_SERVICE_HOURS_DETAILS + " " + month.monthNameEng,
                                month.ajustesSheetName,
                                SHEET_FACTURACION + " " + month.monthNameSpa
                        );
                    }

                    List<String> spanishMonths = availableMonths.stream()
                            .map(m -> m.monthNameSpa)
                            .collect(java.util.stream.Collectors.toList());
                    String outputFileName = fileNameGenerator.generateOutputFileName(spanishMonths, serviceTeam, outputDirectory);
                    Helper.writeWorkbook(outputWorkbook, outputFileName);
                }
            }

            String inputFileName = new File(inputExcelFilePath).getName().replace(".xlsx", "");
            // Create month suffix from all available months
            String monthSuffix = availableMonths.stream()
                    .map(m -> m.monthNameSpa)
                    .collect(java.util.stream.Collectors.joining("_"));
            String consolidatedFileName = outputDirectory
                    + "/Consolidated_Month_Forecast_" + monthSuffix + "_" + inputFileName + ".xlsx";

            try (Workbook consolidatedWorkbook = new XSSFWorkbook()) {
                Map<String, Sheet> monthSheets = new LinkedHashMap<>();
                Map<String, Integer> monthCurrentRows = new LinkedHashMap<>();
                Map<String, List<Integer>> monthGrandTotalHoursRows = new LinkedHashMap<>();
                Map<String, List<Integer>> monthGrandTotalCostRows = new LinkedHashMap<>();
                Map<String, Integer> monthHoursColIndex = new LinkedHashMap<>();
                Map<String, Integer> monthCostColIndex = new LinkedHashMap<>();

                for (MonthSpec month : availableMonths) {
                    String sheetName = "All Teams Forecast " + month.monthNameEng;
                    Sheet monthSheet = consolidatedWorkbook.createSheet(sheetName);
                    monthSheets.put(month.monthNameSpa, monthSheet);
                    monthCurrentRows.put(month.monthNameSpa, 0);
                    monthGrandTotalHoursRows.put(month.monthNameSpa, new ArrayList<>());
                    monthGrandTotalCostRows.put(month.monthNameSpa, new ArrayList<>());
                    int days = Helper.numberOfDays(SHEET_SERVICE_HOURS_DETAILS + " " + month.monthNameEng);
                    monthHoursColIndex.put(month.monthNameSpa, 4 + days);
                    monthCostColIndex.put(month.monthNameSpa, 5 + days);
                }

                for (String serviceTeam : serviceTeamNames) {
                    for (MonthSpec month : availableMonths) {
                        Sheet monthSheet = monthSheets.get(month.monthNameSpa);
                        int currentRow = monthCurrentRows.get(month.monthNameSpa);

                        int nextFreeRow = excelWriter.copyServiceHoursToConsolidatedSheet(
                                inputWorkbook,
                                monthSheet,
                                currentRow,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + month.monthNameSpa,
                                SHEET_SERVICE_HOURS_DETAILS + " " + month.monthNameEng,
                                month.ajustesSheetName,
                                SHEET_FACTURACION + " " + month.monthNameSpa
                        );

                        int totalRowIndex0Based = nextFreeRow - 3;
                        monthGrandTotalHoursRows.get(month.monthNameSpa).add(totalRowIndex0Based);
                        monthGrandTotalCostRows.get(month.monthNameSpa).add(totalRowIndex0Based);
                        monthCurrentRows.put(month.monthNameSpa, nextFreeRow);
                    }
                }

                for (MonthSpec month : availableMonths) {
                    Sheet monthSheet = monthSheets.get(month.monthNameSpa);
                    int currentRow = monthCurrentRows.get(month.monthNameSpa);
                    List<Integer> hoursRows = monthGrandTotalHoursRows.get(month.monthNameSpa);
                    List<Integer> costRows = monthGrandTotalCostRows.get(month.monthNameSpa);

                    if (currentRow > 0) {
                        currentRow += 1;
                        Row grandTotalRow = monthSheet.createRow(currentRow);

                        CellStyle headerStyle = Helper.getHeaderStyle(consolidatedWorkbook);
                        CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(consolidatedWorkbook);
                        int hoursCol = monthHoursColIndex.get(month.monthNameSpa);
                        int costCol = monthCostColIndex.get(month.monthNameSpa);
                        String hoursColLetter = Helper.getColumnLetter(hoursCol);
                        String costColLetter = Helper.getColumnLetter(costCol);

                        Cell labelCell = grandTotalRow.createCell(0);
                        labelCell.setCellValue("GRAND TOTAL (ALL PROJECTS)");
                        labelCell.setCellStyle(headerStyle);
                        monthSheet.addMergedRegion(new CellRangeAddress(currentRow, currentRow, 0, 2));

                        StringJoiner hoursFormula = new StringJoiner("+");
                        for (Integer rowIdx : hoursRows) {
                            hoursFormula.add(hoursColLetter + (rowIdx + 1));
                        }

                        String grandRowNumber = String.valueOf(currentRow + 1);
                        Cell grandRateCell = grandTotalRow.createCell(3);
                        grandRateCell.setCellFormula("IF(" + hoursColLetter + grandRowNumber + "=0,\"\","
                                + costColLetter + grandRowNumber + "/" + hoursColLetter + grandRowNumber + ")");
                        grandRateCell.setCellStyle(footerCurrencyStyle);

                        Cell grandHoursCell = grandTotalRow.createCell(hoursCol);
                        grandHoursCell.setCellFormula(hoursFormula.toString());
                        grandHoursCell.setCellStyle(headerStyle);

                        StringJoiner costFormula = new StringJoiner("+");
                        for (Integer rowIdx : costRows) {
                            costFormula.add(costColLetter + (rowIdx + 1));
                        }

                        Cell grandCostCell = grandTotalRow.createCell(costCol);
                        grandCostCell.setCellFormula(costFormula.toString());
                        grandCostCell.setCellStyle(footerCurrencyStyle);
                    }

                    int lastCol = monthCostColIndex.get(month.monthNameSpa);
                    for (int col = 0; col <= lastCol; col++) {
                        monthSheet.autoSizeColumn(col);
                    }
                }

                Helper.writeWorkbook(consolidatedWorkbook, consolidatedFileName);
            }

        } catch (IOException e) {
            log(logger, "I/O error while processing file '" + file.getName() + "': " + e.getMessage());
        }
    }

    private String toEnglishMonthName(String monthNameSpa) {
        try {
            int monthNumber = getMonthNumber(monthNameSpa, 1);
            return Month.of(monthNumber).getDisplayName(TextStyle.FULL, Locale.US);
        } catch (Exception ignored) {
            if (monthNameSpa == null || monthNameSpa.isBlank()) {
                return monthNameSpa;
            }
            return monthNameSpa.substring(0, 1).toUpperCase(Locale.ROOT) + monthNameSpa.substring(1).toLowerCase(Locale.ROOT);
        }
    }

    private int getMonthNumber(String monthNameSpa, int fallbackMonth) {
        if (monthNameSpa == null) {
            return fallbackMonth;
        }
        String month = monthNameSpa.trim().toLowerCase(Locale.ROOT);
        switch (month) {
            case "enero":
                return 1;
            case "febrero":
                return 2;
            case "marzo":
                return 3;
            case "abril":
                return 4;
            case "mayo":
                return 5;
            case "junio":
                return 6;
            case "julio":
                return 7;
            case "agosto":
                return 8;
            case "septiembre":
                return 9;
            case "octubre":
                return 10;
            case "noviembre":
                return 11;
            case "diciembre":
                return 12;
            default:
                return fallbackMonth;
        }
    }

    private void log(Consumer<String> logger, String message) {
        if (logger != null) {
            logger.accept(message);
        } else {
            System.out.println(message);
        }
    }

    public static void executeScript(String inputExcelFilePath, String outputExcelsFilePath, int monthsToProcess) throws Exception {
        executeScript(inputExcelFilePath, outputExcelsFilePath, monthsToProcess, null);
    }

    public static void executeScript(String inputExcelFilePath, String outputExcelsFilePath, int monthsToProcess, Consumer<String> logger) throws Exception {
        ExecuteService processor = new ExecuteService(
                new DefaultDateProvider(),
                new DefaultExcelFileNameGenerator(),
                new DefaultExcelReader(),
                new ServiceTeamExtractorImpl(),
                new DefaultExcelWriter()
        );
        processor.process(inputExcelFilePath, outputExcelsFilePath, monthsToProcess, logger);
    }
}
