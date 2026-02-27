package invoicing.service.month;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ServiceTeamExtractor {
    List<String> extractFullServiceTeamNames(Sheet sheet, Workbook workbook);

    List<String> extractServiceTeamNames(List<String> fullServiceTeamNames);
}
