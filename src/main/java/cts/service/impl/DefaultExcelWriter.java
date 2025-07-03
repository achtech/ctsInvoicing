package cts.service.impl;


import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.cellwalk.CellHandler;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cts.service.ExcelWriter;
import cts.service.RateTable;
import cts.util.Utils;

public class DefaultExcelWriter implements ExcelWriter {
	private static final Set<Integer> EXCLUDED_ADJUSTMENT_COLUMNS = new HashSet<>(Arrays.asList(0, 1, 2, 7, 8, 10, 11)); // A,

	@Override
	public Workbook createWorkbookWithSheets(String currentMonthName, String nextMonthName, String nextNextMonthName) {
		Workbook workbook = new XSSFWorkbook();
		workbook.createSheet("Service Hours Details " + currentMonthName);
		workbook.createSheet("Service Hours Details " + nextMonthName);
		workbook.createSheet("Service Hours Details " + nextNextMonthName);
//		workbook.createSheet("Adjustment");
//		workbook.createSheet("Invoicing Details " + currentMonthName);
//		workbook.createSheet("Invoicing Details " + nextMonthName);
//		workbook.createSheet("Invoicing Details " + nextNextMonthName);
		return workbook;
	}

	@Override
	public void copyAdjustmentSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam,
			String ajustesSheetName, String adjustmentSheetName) {
		Sheet inputSheet = inputWorkbook.getSheet(ajustesSheetName);
		Sheet outputSheet = outputWorkbook.getSheet(adjustmentSheetName);
		if (inputSheet == null || outputSheet == null) {
			System.err.println("Skipping Adjustment sheet: input or output sheet not found.");
			return;
		}

		int outputRowIndex = 0;

		// Create header style (white font, #003399 background)
		CellStyle currencyStyle = Utils.getCurrencyStyle(outputWorkbook);
		CellStyle headerStyle = Utils.getHeaderStyle(outputWorkbook);
		CellStyle leftStyle = Utils.getLeftStandardStyle(outputWorkbook);
		CellStyle rightStyle = Utils.getRightStandardStyle(outputWorkbook);

		// Create and set custom header row
		Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
		Cell cellA = outputHeaderRow.createCell(0);
		cellA.setCellValue("Client");
		cellA.setCellStyle(headerStyle);

		Cell cellB = outputHeaderRow.createCell(1);
		cellB.setCellValue("Project");
		cellB.setCellStyle(headerStyle);

		Cell cellC = outputHeaderRow.createCell(2);
		cellC.setCellValue("Account");
		cellC.setCellStyle(headerStyle);

		Cell cellD = outputHeaderRow.createCell(3);
		cellD.setCellValue("Description");
		cellD.setCellStyle(headerStyle);

		Cell cellE = outputHeaderRow.createCell(4);
		cellE.setCellValue("Total Hours");
		cellE.setCellStyle(headerStyle);

		Cell cellF = outputHeaderRow.createCell(5);
		cellF.setCellValue("Rates");
		cellF.setCellStyle(headerStyle);

		Cell cellG = outputHeaderRow.createCell(6);
		cellG.setCellValue("Cost");
		cellG.setCellStyle(headerStyle);

		// Copy data rows where column E matches serviceTeam
		for (Row inputRow : inputSheet) {
			if (inputRow.getRowNum() == 0) {
				continue;
			}
			Cell cell = inputRow.getCell(4); // Column E
			if (cell != null && cell.getCellType() == CellType.STRING) {
				if (cell.getStringCellValue().equals(serviceTeam)) {
					Row outputRow = outputSheet.createRow(outputRowIndex++);
					Map<Integer, Integer> columnMapping = new HashMap<>();
					columnMapping.put(3, 0); // D → A
					columnMapping.put(4, 1); // E → B
					columnMapping.put(5, 2); // F → C
					columnMapping.put(6, 3); // G → D
					columnMapping.put(9, 4); // J → E
					columnMapping.put(12, 5); // M → F
					columnMapping.put(13, 6); // N → G
					for (int j = 0; j < inputRow.getLastCellNum(); j++) {
						if (EXCLUDED_ADJUSTMENT_COLUMNS.contains(j)) {
							continue;
						}
						Integer mappedIndex = columnMapping.get(j);
						if (mappedIndex != null) {
//                            System.out.println("mappedIndex:" +mappedIndex);
							Cell inputCell = inputRow.getCell(j);
							Cell outputCell = outputRow.createCell(mappedIndex);
							if (inputCell != null) {
								if (j == 13) { // Column G in input (maps to D in output)
									String formula = "E" + outputRowIndex + "*F" + outputRowIndex; // E*F → B*C
									outputCell.setCellFormula(formula);
									outputCell.setCellStyle(currencyStyle);
								} else {
									switch (inputCell.getCellType()) {
									case NUMERIC:
										outputCell.setCellValue(inputCell.getNumericCellValue());
										outputCell.setCellStyle(rightStyle);

										if (j == 12) {
											outputCell.setCellStyle(currencyStyle);
										}
										break;
									case BOOLEAN:
										outputCell.setCellValue(inputCell.getBooleanCellValue());
										break;
									case FORMULA:
										String formula = "=E" + outputRowIndex + "*F" + outputRowIndex; // E*F → B*C
										outputCell.setCellValue(formula);
										outputCell.setCellStyle(currencyStyle);
										break;
									default:
										outputCell.setCellValue(inputCell.getStringCellValue());
										outputCell.setCellStyle(leftStyle);
										break;
									}
								}
							}
						}
					}
				}
			}
		}
		// Auto-size columns after all data is written
		for (int col = 0; col < 30; col++) { // Adjust up to column G (index 6)
			outputSheet.autoSizeColumn(col);
		}

	}

	private List<Row> getAdjustmentSheetData(Workbook inputWorkbook, String ajustesSheetName, String serviceTeam, int month) {
		Sheet inputSheet = inputWorkbook.getSheet(ajustesSheetName);
		List<Row> rows = new ArrayList<Row>();
		// Copy data rows where column E matches serviceTeam
		for (Row inputRow : inputSheet) {
			if (inputRow.getRowNum() == 0) {
				continue;
			}
			Cell cellE = inputRow.getCell(4); // Column E
			Cell cellH = inputRow.getCell(7); // Column H
			if (cellE.getStringCellValue().equals(serviceTeam) && 
					cellH.getDateCellValue().getMonth() == month-1) {
				rows.add(inputRow);
			}
		}
		return rows;
	}

	@Override
	public void copyFacturationSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam,
			String facturacionSheetName, String invoicingSheetName) {
		Sheet inputSheet = inputWorkbook.getSheet(facturacionSheetName);
		Sheet outputSheet = outputWorkbook.getSheet(invoicingSheetName);
		if (inputSheet == null || outputSheet == null) {
			System.err.println("Skipping invoicing details sheet: input or output sheet not found.");
			return;
		}

		int outputRowIndex = 0;

		// Create header style (white font, #003399 background)
		CellStyle headerStyle = Utils.getHeaderStyle(outputWorkbook);
		CellStyle currencyStyle = Utils.getCurrencyStyle(outputWorkbook);
		CellStyle footerCurrencyStyle = Utils.getFooterCurrencyStyle(outputWorkbook);
		CellStyle centerStandardStyle = Utils.getCenterStandardStyle(outputWorkbook);
		CellStyle leftStandardStyle = Utils.getLeftStandardStyle(outputWorkbook);

		// Create and set custom header row
		Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
		Cell cellA = outputHeaderRow.createCell(0);
		cellA.setCellValue("Employee Number");
		cellA.setCellStyle(headerStyle);

		Cell cellB = outputHeaderRow.createCell(1);
		cellB.setCellValue("Employee Name");
		cellB.setCellStyle(headerStyle);

		Cell cellC = outputHeaderRow.createCell(2);
		cellC.setCellValue("Category");
		cellC.setCellStyle(headerStyle);

		Cell cellD = outputHeaderRow.createCell(3);
		cellD.setCellValue("Rates");
		cellD.setCellStyle(headerStyle);

		Cell cellE = outputHeaderRow.createCell(4);
		cellE.setCellValue("Quantity");
		cellE.setCellStyle(headerStyle);

		Cell cellF = outputHeaderRow.createCell(5);
		cellF.setCellValue("Cost");
		cellF.setCellStyle(headerStyle);

		// Find the row with the service name and copy data until an empty row
		int startCopyRow = -1;
		for (int inputRowIndex = 0; inputRowIndex <= inputSheet.getLastRowNum(); inputRowIndex++) {
			Row inputRow = inputSheet.getRow(inputRowIndex);
			if (inputRow != null) {
				Cell serviceNameCell = inputRow.getCell(1); // Column B (index 1)
				if (serviceNameCell != null && serviceNameCell.getCellType() == CellType.STRING) {
					String serviceName = serviceNameCell.getStringCellValue();
					if (serviceName != null && serviceName.toLowerCase().contains(serviceTeam.toLowerCase())) {
						startCopyRow = inputRowIndex;
						break;
					}
				}
			}
		}

		// If service name is found, copy data from that row until an empty row
		if (startCopyRow != -1) {
			for (int inputRowIndex = startCopyRow + 1; inputRowIndex <= inputSheet.getLastRowNum(); inputRowIndex++) {
				Row inputRow = inputSheet.getRow(inputRowIndex);
				Row outputRow;
				if (inputRow == null || Utils.isRowEmpty(inputRow)) {
					outputRow = outputSheet.createRow(outputRowIndex - 1);
					for (int col = 0; col < 5; col++) { // Adjust up to column G (index 6)
						Cell outputCell = outputRow.createCell(col);
						outputCell.setCellStyle(footerCurrencyStyle);
					}

					Cell outputCellB = outputRow.createCell(1);
					outputCellB.setCellValue("Total : ");
					outputCellB.setCellStyle(headerStyle);

					String formula = "SUM(E2:E" + (outputRowIndex - 1) + ")";
					Cell outputCellCSum = outputRow.createCell(4);
					outputCellCSum.setCellFormula(formula);
					outputCellCSum.setCellStyle(headerStyle);

					String formula2 = "SUM(F2:F" + (outputRowIndex - 1) + ")";
					Cell outputCellESum = outputRow.createCell(5);
					outputCellESum.setCellFormula(formula2);
					outputCellESum.setCellStyle(footerCurrencyStyle);

					break; // Stop copying at an empty row
				}
				outputRow = outputSheet.createRow(outputRowIndex++);

				// Copy Employee Number (Column A)
				Cell outputCellA = outputRow.createCell(0);
				Cell inputCellA = inputRow.getCell(0);
				if (inputCellA != null && inputCellA.getCellType() == CellType.NUMERIC) {
					outputCellA.setCellValue(inputCellA.getNumericCellValue());
					outputCellA.setCellStyle(centerStandardStyle);
				}

				// Copy Employee Name (Column B)
				Cell outputCellB = outputRow.createCell(1);
				Cell inputCellB = inputRow.getCell(1);
				if (inputCellB != null && inputCellB.getCellType() == CellType.STRING) {
					outputCellB.setCellValue(inputCellB.getStringCellValue());
					outputCellB.setCellStyle(leftStandardStyle);
				}

				// Copy Category (Column C)
				Cell outputCellC = outputRow.createCell(2);
				Cell inputCellC = inputRow.getCell(6);
				if (inputCellC != null && inputCellC.getCellType() == CellType.NUMERIC) {
					outputCellC.setCellValue(RateTable.getCategory(inputCellC.getNumericCellValue()));
					outputCellC.setCellStyle(centerStandardStyle);
				}

				// Copy Rate (Column C)
				Cell outputCellD = outputRow.createCell(3);
				Cell inputCellD = inputRow.getCell(6);
				if (inputCellD != null && inputCellD.getCellType() == CellType.NUMERIC) {
					outputCellD.setCellValue(inputCellD.getNumericCellValue());
					outputCellD.setCellStyle(currencyStyle);
				}

				// Copy Category (Column D)
				Cell outputCellE = outputRow.createCell(4);
				Cell inputCellE = inputRow.getCell(3);
				if (inputCellE != null && inputCellE.getCellType() == CellType.NUMERIC) {
					outputCellE.setCellValue(inputCellE.getNumericCellValue());
					outputCellE.setCellStyle(centerStandardStyle);
				}

				// Copy Cost (Column E)
				Cell outputCellF = outputRow.createCell(5);
				Cell inputCellF = inputRow.getCell(7);
				if (inputCellF != null && inputCellF.getCellType() == CellType.FORMULA) {
					String formula = "D" + outputRowIndex + "*E" + outputRowIndex; // E*F → B*C
					outputCellF.setCellFormula(formula);
					outputCellF.setCellStyle(currencyStyle);

				}
			}
		}

		// Auto-size columns after all data is written
		for (int col = 0; col < 20; col++) { // Adjust up to column G (index 6)
			outputSheet.autoSizeColumn(col);
		}
	}

	@Override
	public void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam,
			String facturacionSheetName, String invoicingSheetName, String ajustesSheetName) {
		// TODO Auto-generated method stub
		Sheet inputSheet = inputWorkbook.getSheet(facturacionSheetName);
		Sheet outputSheet = outputWorkbook.getSheet(invoicingSheetName);
		if (inputSheet == null || outputSheet == null) {
			System.err.println("Skipping invoicing details sheet: input or output sheet not found.");
			return;
		}

		int outputRowIndex = 0;

		// Create header style (white font, #003399 background)
		CellStyle headerStyle = Utils.getHeaderStyle(outputWorkbook);
		CellStyle centerStyle = Utils.getCenterStandardStyle(outputWorkbook);
		CellStyle leftStyle = Utils.getLeftStandardStyle(outputWorkbook);
		CellStyle rightStyle = Utils.getRightStandardStyle(outputWorkbook);
		CellStyle currencyStyle = Utils.getCurrencyStyle(outputWorkbook);
		CellStyle dateStyle = Utils.getDateStyle(outputWorkbook);
		CellStyle vacanceStyle = Utils.getVacanceStyle(outputWorkbook);
		CellStyle freedayStyle = Utils.getFreedayStyle(outputWorkbook);
		CellStyle sickLeaveStyle = Utils.getSickLeaveStyle(outputWorkbook);
		CellStyle legalAbsenceStyle = Utils.getLegalAbsenceStyle(outputWorkbook);
		CellStyle weekendStyle = Utils.getWeekendStyle(outputWorkbook);
		CellStyle footerCurrencyStyle = Utils.getFooterCurrencyStyle(outputWorkbook);

		// Create and set custom header row
		Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
		String[] headers = { "Empl. N°", "Person", "Category", "Rates" };
		for (int i = 0; i < headers.length; i++) {
			Cell cell = outputHeaderRow.createCell(i);
			cell.setCellValue(headers[i]);
			cell.setCellStyle(headerStyle);
		}
		int nbrDaysInThisMonths = Utils.numberOfDays(invoicingSheetName);
		for (int i = headers.length, j = 1; i < headers.length + nbrDaysInThisMonths; i++, j++) {
			Cell cell = outputHeaderRow.createCell(i);
			cell.setCellValue(j);
			cell.setCellStyle(dateStyle);
		}
		Cell cell = outputHeaderRow.createCell(headers.length + nbrDaysInThisMonths);
		cell.setCellValue("Working Hours");
		cell.setCellStyle(headerStyle);

		Cell cellAmount = outputHeaderRow.createCell(headers.length + nbrDaysInThisMonths + 1);
		cellAmount.setCellValue("Cost (Euro) ");
		cellAmount.setCellStyle(headerStyle);

		// Iterate through rows in the input sheet
		Map<Double, List<Row>> maps = getAllData(inputSheet, serviceTeam);
		Map<Double, List<Row>> mapsByServiceTeam = filterRowsByServiceTeam(maps, serviceTeam);
		Map<Double, Row> mergedMaps = transformRows(mapsByServiceTeam);
		// WRITE DTAA IN EXCEL FILE
		for (Map.Entry<Double, Row> entry : mergedMaps.entrySet()) {
			Row row = entry.getValue();
			Row outputRow = outputSheet.createRow(outputRowIndex++);
			ALL: for (int j = 0; j < row.getLastCellNum(); j++) {
				Cell inputCell = row.getCell(j);
				Cell outputCell = outputRow.createCell(j);
				if (inputCell != null) {
					switch (inputCell.getCellType()) {
					case NUMERIC:
						outputCell.setCellValue(inputCell.getNumericCellValue());
						outputCell.setCellStyle(centerStyle);
						if (j == 3) {
							outputCell.setCellStyle(currencyStyle);
						}
						break;
					case STRING:
						String value = inputCell.getStringCellValue();
						outputCell.setCellValue(inputCell.getStringCellValue());
						if (j != 1)
							outputCell.setCellStyle("V".equals(value) ? vacanceStyle
									: "A".equals(value) ? legalAbsenceStyle
											: "F".equals(value) ? freedayStyle
													: "S".equals(value) ? sickLeaveStyle : weekendStyle);
						else
							outputCell.setCellStyle(leftStyle);

						break;
						default : 
							break;
					}
				}
				if (j == headers.length + nbrDaysInThisMonths)
					break ALL;
			}
		}

		// Add Cost column
		int lastColumn = outputSheet.getRow(1).getLastCellNum() - 1;
		String letterCost = Utils.getColumnLetter(lastColumn);
		String letterCost2 = Utils.getColumnLetter(lastColumn+1);
		String letterTotalHours = Utils.getColumnLetter(lastColumn);
		
		for (Row row : outputSheet) {
			if (row.getRowNum() != 0) {
				// Create a new cell in the last column (new column)
				Cell newCell = row.createCell(lastColumn + 1);
				// Create formula: D(row_number) * last_column
				String formula = "D" + (row.getRowNum() + 1) + "*" + letterCost	+ (row.getRowNum() + 1);
				newCell.setCellFormula(formula);
				newCell.setCellStyle(currencyStyle);
			}
		}
		// ADD ADJUSTMENT
		int month = Utils.getMonthFromSheetName(invoicingSheetName);
		List<Row> AjustesRows = getAdjustmentSheetData(inputWorkbook, ajustesSheetName, serviceTeam,month );
		for (Row row : AjustesRows) {
			Row outputRow = outputSheet.createRow(outputRowIndex++);

			Cell cellA = outputRow.createCell(0);
			cellA.setCellValue("");
			cellA.setCellStyle(centerStyle);
			
			Cell cellB = outputRow.createCell(1);
			cellB.setCellValue(row.getCell(6).getStringCellValue());
			cellB.setCellStyle(leftStyle);
			
			Cell cellC = outputRow.createCell(2);
			cellC.setCellValue("");
			cellC.setCellStyle(centerStyle);
			
			Cell cellD = outputRow.createCell(3);
			cellD.setCellValue(row.getCell(12).getNumericCellValue());
			cellD.setCellStyle(currencyStyle);
			
			Cell cellE = outputRow.createCell(4);
			cellE.setCellValue(row.getCell(9).getNumericCellValue());
			cellE.setCellStyle(centerStyle);
			int lastCol = headers.length + nbrDaysInThisMonths;
			for (int i = headers.length+1; i < lastCol; i++) {
				Cell cellM = outputRow.createCell(i);
				cellM.setCellValue("");
				cellM.setCellStyle(centerStyle);
			}
			Cell cellAdj = outputRow.createCell(lastCol);
			cellAdj.setCellValue(row.getCell(9).getNumericCellValue());
			cellAdj.setCellStyle(centerStyle);
			
			Cell cellCost= outputRow.createCell(lastCol+1);
			String formulaCost = "D"+outputRowIndex+"*E"+outputRowIndex;
			cellCost.setCellFormula(formulaCost);
			cellCost.setCellStyle(currencyStyle);
			
		}		
		// ADD TOTAL ROW
		Row lastRow = outputSheet.createRow(outputRowIndex);
		Cell cellTotal = lastRow.createCell(lastColumn - 2);
		cellTotal.setCellValue("Total");
		cellTotal.setCellStyle(headerStyle);
		outputSheet.addMergedRegion(new CellRangeAddress(
				outputRowIndex, // First row (0-based)
				outputRowIndex, // Last row (0-based)
				lastColumn - 2, // First column (0-based)
				lastColumn - 1  // Last column (0-based)
	        ));

		Cell cellTotalHours = lastRow.createCell(lastColumn );
		String formula = "SUM("+letterTotalHours+"2:"+letterTotalHours+(outputRowIndex)+")";
		cellTotalHours.setCellFormula(formula);
		cellTotalHours.setCellStyle(headerStyle);

		Cell cellTotalCost = lastRow.createCell(lastColumn+1 );
		String formula2 = "SUM("+letterCost2+"2:"+letterCost2+(outputRowIndex)+")";
		cellTotalCost.setCellFormula(formula2);
		cellTotalCost.setCellStyle(footerCurrencyStyle);
		
		// Auto-size columns after all data is written
		for (int col = 0; col < 40; col++) { // Adjust up to column G (index 6)
			outputSheet.autoSizeColumn(col);
		}

	}

	private Map<Double, List<Row>> getAllData(Sheet inputSheet, String serviceTeam) {
		List<Double> ids = new ArrayList<Double>();
		Map<Double, List<Row>> maps = new HashMap<Double, List<Row>>();
		Double lastId = 0D;
		for (Row row : inputSheet) {
			Cell empIdCell = row.getCell(0);
			Cell empNameCell = row.getCell(1);
			if (empIdCell != null && empNameCell != null) {
				lastId = empIdCell.getNumericCellValue();
				ids.add(lastId);
				List<Row> list = new ArrayList<Row>();
				list.add(row);
				maps.put(empIdCell.getNumericCellValue(), list);
			} else {
				if (!Utils.isRowEmpty(row) && lastId != 0) {
					List<Row> list = maps.get(lastId);
					list.add(row);
					maps.put(lastId, list);
				}
			}
		}
		return maps;
	}

	private static Map<Double, List<Row>> filterRowsByServiceTeam(Map<Double, List<Row>> inputMap, String serviceTeam) {
		Map<Double, List<Row>> filteredMap = new HashMap<>();

		for (Map.Entry<Double, List<Row>> entry : inputMap.entrySet()) {
			Double key = entry.getKey();
			List<Row> rows = entry.getValue();

			// Extract the first row and filter rows based on the second cell condition
			List<Row> filteredRows = rows.stream().filter(row -> {
				Cell secondCell = row.getCell(1); // Second cell (index 1)
				boolean isFirstRow = rows.indexOf(row) == 0; // Check if it's the first row
				boolean isSecondCellEmptyOrMatches = secondCell == null || secondCell.getCellType() == CellType.BLANK
						|| (secondCell.getCellType() == CellType.STRING
								&& secondCell.getStringCellValue().contains(serviceTeam));
				return isFirstRow || isSecondCellEmptyOrMatches;
			}).collect(Collectors.toList());

			if (!filteredRows.isEmpty()) {
				filteredMap.put(key, filteredRows);
			}
		}

		return filteredMap;
	}

	private Map<Double, Row> transformRows(Map<Double, List<Row>> inputMap) {
		Map<Double, Row> resultMap = new HashMap<>();
		for (Map.Entry<Double, List<Row>> entry : inputMap.entrySet()) {
			Double key = entry.getKey();
			List<Row> rows = entry.getValue();

			if (rows == null || rows.isEmpty()) {
				continue; // Skip if the list is empty or null
			}

			// Create a new row for the result
			Workbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
			Row newRow = workbook.createSheet().createRow(0);

			// Cell 0: First cell of the first row
			if (!rows.isEmpty()) {
				Cell firstCell = rows.get(0).getCell(0);
				newRow.createCell(0).setCellValue(firstCell != null ? firstCell.getNumericCellValue() : 0);
			}

			// Cell 1: Second cell of the first row
			if (!rows.isEmpty() && rows.get(0).getCell(1) != null) {
				Cell secondCellFirst = rows.get(0).getCell(1);
				newRow.createCell(1).setCellValue(secondCellFirst.getStringCellValue());
			}

			// Cell 1: Second cell of the first row
			if (!rows.isEmpty() && rows.size() > 1 && rows.get(1) != null && rows.get(1).getCell(1) != null) {
				Cell secondCellSecond = rows.get(1).getCell(1);
				Double input = Utils.getRates(secondCellSecond.getStringCellValue());

				newRow.createCell(2).setCellValue(RateTable.getCategory(input));
			}

			// Cell 2: Second cell of the second row
			CellStyle currencyStyle = Utils.getCurrencyStyle(workbook);
			if (rows.size() > 1 && rows.get(1).getCell(1) != null) {
				Cell secondCellSecond = rows.get(1).getCell(1);
				Double input = Utils.getRates(secondCellSecond.getStringCellValue());
				Cell thirdCell = newRow.createCell(3);
				thirdCell.setCellValue(input);
				thirdCell.setCellStyle(currencyStyle);
			}

			// Cells N (N >= 3): Based on the third cell and beyond
			if (rows.size() > 1) {
				Row teamServiceRow = rows.get(1);
				Row vacationRow = rows.size() > 2 ? rows.get(2) : null;
				for (int i = 4; i < teamServiceRow.getLastCellNum() + 4; i++) { // Limit to 10 cells for practicality,
																				// adjust as needed
					Cell hoursCell = teamServiceRow.getCell(i - 2);
					Cell vacationsCell = vacationRow != null ? vacationRow.getCell(i - 2) : null;
					Cell outputCell = newRow.createCell(i);
					if (hoursCell != null && hoursCell.getCellType() != CellType.BLANK) {
						switch (hoursCell.getCellType()) {
						case NUMERIC:
							outputCell.setCellValue(hoursCell.getNumericCellValue());
							break;
						case FORMULA:
							Workbook w= teamServiceRow.getSheet().getWorkbook();
							FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator();

					        // Evaluate the formula cell to get its numeric value
					        CellValue cellValue = evaluator.evaluate(hoursCell);
					        double numericValue = cellValue.getNumberValue();
							outputCell.setCellValue(numericValue);
							break;
						default:
							outputCell.setCellValue(hoursCell.getStringCellValue());
							break;
						}
					}
					if (outputCell == null || outputCell.getCellType() == CellType.BLANK
							|| (outputCell.getCellType() == CellType.STRING
									&& outputCell.getStringCellValue().isEmpty())) {
						if (vacationsCell != null) {
							switch (vacationsCell.getCellType()) {
							case NUMERIC:
								outputCell.setCellValue(vacationsCell.getNumericCellValue());
								break;
							case STRING:
								outputCell.setCellValue(!vacationsCell.getStringCellValue().isEmpty()
										? vacationsCell.getStringCellValue()
										: "");
								break;
							default:
								break;
							}
						}

					}
				}
			}

			if (newRow.getCell(2) != null)
				resultMap.put(key, newRow);
		}

		return resultMap;
	}
}