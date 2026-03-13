package invoicing.service.month;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.math.BigDecimal;
import java.util.List;

public interface ExcelWriter {
    Workbook createWorkbookWithSheets(List<String> monthNames);
    BigDecimal getTotalServiceTeam(Workbook inputWorkbook, String serviceTeam, String sheetName);
    void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam, String HorasServicioSheetName, String ServiceHoursSheetName, String ajustesSheetName, String facturacionSheetName);
    int copyServiceHoursToConsolidatedSheet(
            Workbook inputWorkbook,
            Sheet    consolidatedSheet,
            int      startRow,
            String   serviceTeam,
            String   invoicingSheetNameES,
            String   invoicingSheetNameEN,
            String   ajustesSheetName,
            String   facturacionSheetName
    );
}
