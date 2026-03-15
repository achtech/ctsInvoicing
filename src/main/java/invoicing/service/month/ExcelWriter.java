package invoicing.service.month;

import org.apache.poi.ss.usermodel.Workbook;
import java.util.List;

public interface ExcelWriter {
    Workbook createWorkbookWithSheets(List<String> monthNames);
    void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String HorasServicioSheetName, String ServiceHoursSheetName, String ajustesSheetName, String facturacionSheetName);
}
