package invoicing.service.month.impl;

import invoicing.service.month.ServiceTeamExtractor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.ArrayList;
import java.util.List;

public class ServiceTeamExtractorImpl implements ServiceTeamExtractor {
    @Override
    public List<String> extractFullServiceTeamNames(Sheet sheet, Workbook workbook) {
        List<String> fullServiceTeamsName = new ArrayList<>();
        for (Row row : sheet) {
            Cell cell = row.getCell(1); // Column B
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String value = cell.getStringCellValue();
                if (value != null && !value.isEmpty() && value.contains(">")) {
                    fullServiceTeamsName.add(value);
                }
            }
        }
        return fullServiceTeamsName;
    }

    @Override
    public List<String> extractServiceTeamNames(List<String> fullServiceTeamNames) {
        List<String> serviceTeamsName = new ArrayList<>();
        for (String fullName : fullServiceTeamNames) {
            String[] parts = fullName.split(">");
            if (parts.length > 1) {
                serviceTeamsName.add(parts[1].trim());
            }
        }
        return serviceTeamsName;
    }
}
