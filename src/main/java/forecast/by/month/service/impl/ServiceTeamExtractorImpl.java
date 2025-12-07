package forecast.by.month.service.impl;

import forecast.by.month.service.ServiceTeamExtractor;
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
                CellStyle style = cell.getCellStyle();
                if (style != null) {
                    Font font = workbook.getFontAt(style.getFontIndexAsInt());
                    if (font instanceof XSSFFont) {
                        XSSFFont xssfFont = (XSSFFont) font;
                        XSSFColor color = xssfFont.getXSSFColor();
                        if (color != null && color.getRGB() != null) {
                            byte[] rgb = color.getRGB();
                            if (rgb[0] == (byte) 255 && rgb[1] == (byte) 255 && rgb[2] == (byte) 255) {
                                fullServiceTeamsName.add(cell.getStringCellValue());
                            }
                        }
                    }
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
