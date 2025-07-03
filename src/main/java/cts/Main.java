package cts;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cts.service.*;
import cts.service.impl.*;
import cts.util.Utils;

import java.io.*;
import java.time.LocalDate;
import java.util.*;
import static cts.service.ExcelFileNameGenerator.*;

public class Main {
	private final DateProvider dateProvider;
    private final ExcelFileNameGenerator fileNameGenerator;
    private final ExcelReader excelReader;
    private final ServiceTeamExtractor serviceTeamExtractor;
    private final ExcelWriter excelWriter;

    public Main(
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

    public void process() {
    	Scanner scanner = new Scanner(System.in);
        System.out.println("Enter the Excel file path (e.g., C:\\Temp\\data.xlsx):");
        String filePath = scanner.nextLine().trim(); // Trim to remove extra spaces
        scanner.close();

        // Normalize the path for cross-platform compatibility
        File file = new File(filePath).getAbsoluteFile();
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


            LocalDate currentDate = dateProvider.getCurrentDate(filePath);
            int currentYear = dateProvider.getYear(currentDate);
            int currentMonth = dateProvider.getMonthValue(currentDate);
            String currentMonthName = dateProvider.getMonthNameEnglish(currentDate);
            String nextMonthName = dateProvider.getMonthNameEnglish(currentDate.plusMonths(1));
            String nextNextMonthName = dateProvider.getMonthNameEnglish(currentDate.plusMonths(2));
            String currentMonthSpanish = dateProvider.getMonthNameSpanish(currentDate);
            String nextMonthSpanish = dateProvider.getMonthNameSpanish(currentDate.plusMonths(1));
            String nextNextMonthSpanish = dateProvider.getMonthNameSpanish(currentDate.plusMonths(2));

//            LocalDate currentDate = dateProvider.getCurrentDate();
//            int currentYear = dateProvider.getYear(currentDate);
//            int currentMonth = dateProvider.getMonthValue(currentDate);
//            String currentMonthName = dateProvider.getMonthNameEnglish(currentDate);
//            String nextMonthName = dateProvider.getMonthNameEnglish(currentDate.plusMonths(1));
//            String nextNextMonthName = dateProvider.getMonthNameEnglish(currentDate.plusMonths(2));
//            String currentMonthSpanish = dateProvider.getMonthNameSpanish(currentDate);
//            String nextMonthSpanish = dateProvider.getMonthNameSpanish(currentDate.plusMonths(1));
//            String nextNextMonthSpanish = dateProvider.getMonthNameSpanish(currentDate.plusMonths(2));
            String outputDirectory = Utils.getDesktopPath()+ "/INVOICING/" + currentMonthName + "_" + currentYear;

            Sheet facturacionSheet = excelReader.getSheet(inputWorkbook, "Facturación " + currentMonthSpanish);

            List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
            List<String> serviceTeamNames = serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames);

            for (String serviceTeam : serviceTeamNames) {
                try (Workbook outputWorkbook = excelWriter.createWorkbookWithSheets(currentMonthName, nextMonthName, nextNextMonthName)) {
//                    excelWriter.copyAdjustmentSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_AJUSTES, SHEET_ADJUSTMENT);
//                    excelWriter.copyFacturationSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_FACTURACION+" "+currentMonthSpanish, SHEET_INVOICING_DETAILS+" "+currentMonthName);
//                    excelWriter.copyFacturationSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_FACTURACION+" "+nextMonthSpanish, SHEET_INVOICING_DETAILS+" "+nextMonthName);
//                    excelWriter.copyFacturationSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_FACTURACION+" "+nextNextMonthSpanish, SHEET_INVOICING_DETAILS+" "+nextNextMonthName);
                    excelWriter.copyServiceHoursSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_HORAS_SERVICIO+" "+currentMonthSpanish, SHEET_SERVICE_HOURS_DETAILS+" "+currentMonthName,SHEET_AJUSTES);
                    excelWriter.copyServiceHoursSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_HORAS_SERVICIO+" "+nextMonthSpanish, SHEET_SERVICE_HOURS_DETAILS+" "+nextMonthName,SHEET_AJUSTES);
                    excelWriter.copyServiceHoursSheetData(inputWorkbook, outputWorkbook, serviceTeam, SHEET_HORAS_SERVICIO+" "+nextNextMonthSpanish, SHEET_SERVICE_HOURS_DETAILS+" "+nextNextMonthName,SHEET_AJUSTES);
                    String outputFileName = fileNameGenerator.generateOutputFileName(currentMonth, currentYear, serviceTeam, outputDirectory);
                    Utils.writeWorkbook(outputWorkbook, outputFileName);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Main processor = new Main(
                new DefaultDateProvider(),
                new DefaultExcelFileNameGenerator(),
                new DefaultExcelReader(),
                new ServiceTeamExtractorImpl(),
                new DefaultExcelWriter()
        );
        processor.process();
    }
}