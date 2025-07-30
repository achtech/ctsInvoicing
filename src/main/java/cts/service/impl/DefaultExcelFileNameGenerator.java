package cts.service.impl;

import cts.service.ExcelFileNameGenerator;

import java.io.File;

public class DefaultExcelFileNameGenerator implements ExcelFileNameGenerator {
    @Override
    public String generateOutputFileName(int month, int year, String serviceTeam, String directory) {
        return directory + File.separator + month + "_" + year + "_" + serviceTeam + ".xlsx";
    }
}