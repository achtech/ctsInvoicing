package invoicing.service.month.impl;



import java.io.File;
import java.util.List;
import java.util.stream.Collectors;

import invoicing.service.month.ExcelFileNameGenerator;

public class DefaultExcelFileNameGenerator implements ExcelFileNameGenerator {
    @Override
    public String generateOutputFileName(int month, int year, String serviceTeam, String directory) {
        return directory + File.separator + month + "_" + year + "_" + serviceTeam + ".xlsx";
    }
    
    @Override
    public String generateOutputFileName(List<String> months, String serviceTeam, String directory) {
        String monthStr = months.stream().collect(Collectors.joining("_"));
        return directory + File.separator + monthStr + "_" + serviceTeam + ".xlsx";
    }
}