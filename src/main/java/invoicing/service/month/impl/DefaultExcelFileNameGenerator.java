package invoicing.service.month.impl;



import java.io.File;

import invoicing.service.month.ExcelFileNameGenerator;

public class DefaultExcelFileNameGenerator implements ExcelFileNameGenerator {
    @Override
    public String generateOutputFileName(int month, int year, String serviceTeam, String directory) {
        return directory + File.separator + month + "_" + year + "_" + serviceTeam + ".xlsx";
    }
}