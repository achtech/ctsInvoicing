package cts.service;

import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelWriter {
    Workbook createWorkbookWithSheets(String currentMonthName, String nextMonthName, String nextNextMonthName);

    void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String HorasServicioSheetName, String ServiceHoursSheetName, String ajustesSheetName);
}
