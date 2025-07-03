package cts.service;

import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelWriter {
    Workbook createWorkbookWithSheets(String currentMonthName, String nextMonthName, String nextNextMonthName);
    void copyAdjustmentSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String ajustesSheetName, String adjustmentSheetName);
    void copyFacturationSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String FacturacionSheetName, String invoicingSheetName);
    void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String HorasServicioSheetName, String ServiceHoursSheetName, String ajustesSheetName);
}
