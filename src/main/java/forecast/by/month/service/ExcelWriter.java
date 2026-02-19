package forecast.by.month.service;

import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;
import java.util.List;

public interface ExcelWriter {
    Workbook createWorkbookWithSheets(List<String> monthNames);
    BigDecimal getTotalServiceTeam(Workbook inputWorkbook, String serviceTeam, String sheetName);
    void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String HorasServicioSheetName, String ServiceHoursSheetName, String ajustesSheetName, String facturacionSheetName);
}
