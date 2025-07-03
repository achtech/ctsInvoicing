package cts.service.impl;

import java.io.File;

import cts.service.ExcelFileNameGenerator;

public class DefaultExcelFileNameGenerator implements ExcelFileNameGenerator {
	    @Override
	    public String generateInputFileName(int year, int month) {
	        return year + "-" + month + "-Intesa_Nautilus.xlsx";
	    }

	    @Override
	    public String generateOutputFileName(int month, int year, String serviceTeam, String directory) {
	        return directory + File.separator + month + "_" + year + "_" + serviceTeam + ".xlsx";
	    }
	}