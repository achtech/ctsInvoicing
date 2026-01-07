package forecast.by.month.service;

import forecast.by.month.service.impl.*;
import forecast.by.month.util.Utils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;

import static forecast.by.month.service.ExcelFileNameGenerator.*;

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

    public void process(String inputExcelFilePath, String outputGeneratedExcelsFilePath) {
        // Normalize the path for cross-platform compatibility
        File file = new File(inputExcelFilePath).getAbsoluteFile();
        System.out.println("Resolved path: " + file.getAbsolutePath());

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
        try (FileInputStream fis = new FileInputStream(file);
             Workbook inputWorkbook = new XSSFWorkbook(fis)) {


            LocalDate currentDate = dateProvider.getCurrentDate(inputExcelFilePath);
            int currentYear = dateProvider.getYear(currentDate);
            int currentMonth = dateProvider.getMonthValue(currentDate);
            String currentMonthName = dateProvider.getMonthNameEnglish(currentDate);
            String nextMonthName = dateProvider.getMonthNameEnglish(currentDate.plusMonths(1));
            String nextNextMonthName = dateProvider.getMonthNameEnglish(currentDate.plusMonths(2));
            String currentMonthSpanish = dateProvider.getMonthNameSpanish(currentDate);
            String nextMonthSpanish = dateProvider.getMonthNameSpanish(currentDate.plusMonths(1));
            String nextNextMonthSpanish = dateProvider.getMonthNameSpanish(currentDate.plusMonths(2));

            String outputDirectory = outputGeneratedExcelsFilePath + "/Version_" + dateProvider.getCurrentDateTime();

            Sheet facturacionSheet = excelReader.getSheet(inputWorkbook, "Facturación " + currentMonthSpanish);

            List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
            List<String> serviceTeamNames = serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames);

            for (String serviceTeam : serviceTeamNames) {
                try (Workbook outputWorkbook = excelWriter.createWorkbookWithSheets(currentMonthName, nextMonthName, nextNextMonthName)) {
                    excelWriter.copyServiceHoursSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_HORAS_SERVICIO + " " + currentMonthSpanish, SHEET_SERVICE_HOURS_DETAILS + " " + currentMonthName, SHEET_AJUSTES, SHEET_FACTURACIÓN + " " + currentMonthSpanish);
                    excelWriter.copyServiceHoursSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_HORAS_SERVICIO + " " + nextMonthSpanish, SHEET_SERVICE_HOURS_DETAILS + " " + nextMonthName, SHEET_AJUSTES, SHEET_FACTURACIÓN + " " + nextMonthSpanish);
                    excelWriter.copyServiceHoursSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_HORAS_SERVICIO + " " + nextNextMonthSpanish, SHEET_SERVICE_HOURS_DETAILS + " " + nextNextMonthName, SHEET_AJUSTES, SHEET_FACTURACIÓN + " " + nextNextMonthSpanish);
                    String outputFileName = fileNameGenerator.generateOutputFileName(currentMonth, currentYear, serviceTeam, outputDirectory);
                    Utils.writeWorkbook(outputWorkbook, outputFileName);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void executeScript(String inputExcelFilePath, String outputExcelsFilePath) {
        ExecuteService processor = new ExecuteService(
                new DefaultDateProvider(),
                new DefaultExcelFileNameGenerator(),
                new DefaultExcelReader(),
                new ServiceTeamExtractorImpl(),
                new DefaultExcelWriter()
        );
        processor.process(inputExcelFilePath, outputExcelsFilePath);
    }
}