package cts.service;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface ServiceTeamExtractor {
    List<String> extractFullServiceTeamNames(Sheet sheet, Workbook workbook);
    List<String> extractServiceTeamNames(List<String> fullServiceTeamNames);
}
