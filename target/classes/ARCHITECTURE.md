# CTS Invoicing – Architecture Overview

This document describes the current architecture of the CTS Invoicing desktop application, the purpose of each package, and the role of every Java class. It also explains how data flows through the system and how the different parts collaborate.

---

## 1. High‑Level Architecture

- Desktop application built with Swing.
- Single main window (`InvoicingDashboard`) where the user:
  - Chooses input Excel files.
  - Chooses an output directory.
  - Configures the number of months for the Month module.
  - Runs all modules and views logs.
- Application services and helpers implement three main processing modules:
  - Rate module.
  - ExtCode module.
  - Month module.
- A central orchestrator (`UnifiedExecutionService`) coordinates these three modules when the user presses “RUN ALL PROCESSES”.

Execution flow:

1. `invoicing.Main` starts the Swing application and opens `InvoicingDashboard`.
2. `InvoicingDashboard.AllInOnePanel` collects user input.
3. It calls `UnifiedExecutionService.runUnified(...)` with:
   - Target directory.
   - List of input Excel files.
   - Number of months and “manual/auto” flag.
   - A listener for logs and progress.
4. `UnifiedExecutionService`:
   - Creates a timestamped output root directory.
   - Runs the Rate module.
   - Runs the ExtCode module.
   - Runs the Month module.
5. Each module uses helpers and entities to read Excel, process data, and write new Excel files.

---

## 2. Package Layout

Root source directory:

- `invoicing`
  - `Main.java`
  - `Helper/`
  - `entities/`
  - `enums/`
  - `service/`
    - `ext/`
    - `month/`
      - `impl/`
    - `rate/`
    - `UnifiedExecutionService.java`
  - `view/`
    - `InvoicingDashboard.java`
    - `UnifiedMain.java` (legacy UI)

Resources:

- `src/main/resources/Data.xlsx`
- `src/main/resources/history.csv`
- `src/main/resources/README`
- `src/main/resources/ARCHITECTURE.md` (this document)

---

## 3. Application Entry Point and UI Layer

### 3.1 `invoicing.Main`

- Location: `src/main/java/invoicing/Main.java`
- Responsibilities:
  - Configures Swing Look & Feel (Nimbus or system).
  - Sets global fonts through `UIManager`.
  - Starts the main window:
    - `SwingUtilities.invokeLater(() -> new InvoicingDashboard().setVisible(true));`
- This is the only `main` method intended to start the application.

### 3.2 `invoicing.view.InvoicingDashboard`

- Location: `src/main/java/invoicing/view/InvoicingDashboard.java`
- Responsibilities:
  - Main Swing window for the tool (“CTS Invoicing Dashboard”).
  - Contains the nested static class `AllInOnePanel`, which builds the full UI:
    - File list for input Excel files.
    - Target output directory selection.
    - Forecast months spinner and toggle (manual/auto).
    - “RUN ALL PROCESSES” button.
    - Progress bar and status label.
    - Log area where backend messages are displayed.
  - Uses `HISTORY_PATH = "src/main/resources/history.csv"` to remember the last output directory and month count between runs.

Key interactions:

- When the user presses “RUN ALL PROCESSES”:
  - Validates that there is at least one input file.
  - Validates that the target directory exists and is a directory.
  - Reads months and `useManual` from the UI.
  - Saves history to `history.csv`.
  - Builds a `List<File>` from the JList model.
  - Disables buttons and resets progress.
  - Creates a `UnifiedExecutionService` and a `Listener` implementation that:
    - Forwards `log(message)` to the text area.
    - Forwards `setProgress(value, label, detail)` to the progress bar and status label.
  - Starts a background thread that calls:
    - `service.runUnified(targetDir, inputs, months, useManual, listener)`.
  - When the service returns, it:
    - Re‑enables the run button.
    - Updates `openOutputBtn` to point at the last main output folder.
    - Shows a message dialog with the final output path.

### 3.3 `invoicing.view.UnifiedMain` (legacy UI)

- Location: `src/main/java/invoicing/view/UnifiedMain.java`
- Legacy version of the UI that still contains embedded logic similar to what is now in `UnifiedExecutionService`.
- Keeps its own `main` method and `AllInOnePanel` implementation.
- Uses the same history file path `src/main/resources/history.csv`.
- The recommended entry point now is `invoicing.Main` plus `InvoicingDashboard`.

---

## 4. Orchestration Layer

### 4.1 `invoicing.service.UnifiedExecutionService`

- Location: `src/main/java/invoicing/service/UnifiedExecutionService.java`
- Responsibilities:
  - Orchestrates the three main processing modules: Rate, ExtCode, and Month.
  - Coordinates output folder creation and logging.

Key API:

- `public interface Listener`
  - `void log(String message);`
  - `void setProgress(int value, String barLabel, String detail);`
  - Implemented by the UI to display progress and logs.

- `public File runUnified(File targetDir, List<File> inputs, int months, boolean useManual, Listener listener)`
  - Creates the root output folder:
    - Name: `forecast_italy_<MMM_yyyy>_<yyyyMMdd_HHmmss>`.
  - Creates subfolders:
    - `forecast_it_rate_<MMM_yyyy>` for Rate module.
    - `forecast_EXT_<MMM_yyyy>` for ExtCode module.
    - `forecast_month_<MMM_yyyy>` for Month module.
  - Logs start messages and months setting.
  - Calls:
    - `runRateModule(...)`
    - `runExtModule(...)`
    - `runMonthModule(...)`
  - When done:
    - Sets progress to 3 (“Completed”).
    - Logs completion.
    - Returns the root output folder.

Internal methods:

- `runRateModule(LocalDateTime now, File rateFolder, List<File> inputs, Listener listener)`
  - Loads `Data.xlsx` via `ReferenceData`.
  - Builds `GroupAggregator` and `InputRowProcessor`.
  - Uses `InputFilesReader` to process each input Excel file.
  - Aggregates hours and cost per GroupId.
  - Writes a consolidated “Rate Forecast <Month>.xlsx” into `rateFolder` via `OutputWriter`.

- `runExtModule(File extFolder, List<File> inputs, Listener listener)`
  - Uses `service.ext.ExcelReader` to extract raw service team labels and costs from each input Excel file.
  - Uses `ServiceTeamParser` to convert labels into `ServiceTeam` entities:
    - Extracts BU.
    - Extracts EXT code.
    - Extracts project name and descriptions.
  - Copies cost and cell style from raw rows into `ServiceTeam`.
  - Uses `service.ext.ExcelWriter` to create a forecast Excel file in `extFolder`.

- `runMonthModule(File monthFolder, List<File> inputs, int months, boolean useManual, Listener listener)`
  - For each input file:
    - If `useManual` is false, calls `countMonthSheets(...)` to detect how many “Facturacion” sheets exist.
    - Logs detection or warnings if none found.
    - Calls `ExecuteService.executeScript(inputPath, monthFolderPath, currentMonths)`.
  - Logs that month processing finished.

- `countMonthSheets(File f, Listener listener)`
  - Uses Apache POI to open the workbook and count sheets whose normalized name contains “facturacion”.
  - Logs any errors via `listener`.

---

## 5. Rate Module (`invoicing.service.rate`)

This module calculates a consolidated Excel report aggregated by GroupId using the “rate” logic.

### 5.1 `InputFilesReader`

- Location: `src/main/java/invoicing/service/rate/InputFilesReader.java`
- Responsibilities:
  - Open each input Excel file.
  - Locate the relevant sheet for the current month:
    - Searches for sheet names containing “Facturación” and the current month (in Spanish).
  - Iterate rows and delegate to `InputRowProcessor`.
  - Add results to `GroupAggregator`.

Key logic:

- `processFile(String filePath)`
  - For each row (skipping header):
    - Calls `rowProcessor.processRow(row)` to extract `RowData`.
    - If non‑null, calls `aggregator.add(groupId, "user", hours, cost)`.

### 5.2 `InputRowProcessor`

- Location: `src/main/java/invoicing/service/rate/InputRowProcessor.java`
- Responsibilities:
  - Interpret one row of the rate input sheet.
  - Extract:
    - Rate information from text columns.
    - Hours, invoiced rate, and cost.
  - Correct inconsistencies between invoice rate and reference rate using `ReferenceData`.
  - Return a normalized `RowData` with:
    - GroupId.
    - Hours.
    - Cost.

Key steps in `processRow(Row row)`:

1. Extract text and numeric values from specific columns (B, C, D, G, H).
2. Call `extractRateFromText(...)` to parse “EUR” rate from column text.
3. If rate is zero, fall back to other cells when possible.
4. Use `ReferenceData` to:
   - Find approximate GroupId.
   - Find the correct rate.
5. Adjust hours and cost depending on comparison between invoice rate and correct rate:
   - Special guardias cases (25 and 50).
   - General case adjusts hours and recalculates cost.
6. Add a `ProcessedDataEntry` to an internal list for debugging.
7. Return `RowData(groupId, hours, cost)` or `null` if group cannot be resolved.

### 5.3 `OutputWriter`

- Location: `src/main/java/invoicing/service/rate/OutputWriter.java`
- Responsibilities:
  - Produce a consolidated Excel report by GroupId using:
    - `ReferenceData` for rate per group.
    - `GroupAggregator` for hours and cost per group.

Key logic:

- `write(String outputPath)`
  - Creates a workbook and a “Consolidated” sheet.
  - Writes header row: Client, Project, GroupId, Rate, Hours, Cost.
  - For each GroupId:
    - Computes total hours and total cost from `GroupAggregator`.
    - Writes a data row:
      - Client: “Italy”.
      - Project: “INS-026696-00003”.
      - GroupId, rate, hours, cost.
  - Writes a final total row with the sum of all costs.
  - Sets column widths and saves to `outputPath`.

---

## 6. ExtCode Module (`invoicing.service.ext`)

This module reads external code information from Excel and produces a summary forecast by service team.

### 6.1 `ExcelReader`

- Location: `src/main/java/invoicing/service/ext/ExcelReader.java`
- Responsibilities:
  - Read input Excel files and extract “service team” rows and their costs.

Key logic:

- Inner class `ServiceTeamRaw`
  - Fields: `label`, `cost`, `style`.
  - Represents a raw entry from the input sheet.

- `extractRawServiceTeams(File file)`
  - Opens the workbook and finds the appropriate sheet:
    - Looks for “Facturación” plus current Spanish month name.
  - Scans rows:
    - Detects label rows in column B with a non‑null fill color.
    - Tracks current label and last numeric cost in column H.
    - When encountering an empty row, closes the current group and adds a `ServiceTeamRaw`.
  - Returns a list of `ServiceTeamRaw` entries, removing empty labels.

### 6.2 `ServiceTeamParser`

- Location: `src/main/java/invoicing/service/ext/ServiceTeamParser.java`
- Responsibilities:
  - Convert raw service team label strings into rich `ServiceTeam` entities.
  - Use mapping tables from `ServiceTeamMaps`.

Key logic:

- `parse(List<String> rawItems)`
  - For each raw label:
    - Splits by `">"` and takes the second part when available.
    - `extractBU(...)`:
      - Matches the beginning of the string with keys from `BU_DESCRIPTION_MAP`.
    - `extractEXT(...)`:
      - Finds the substring starting with “EXT”.
      - Default “Pending” if none found.
    - `extractProjectName(...)`:
      - Uses BU and EXT positions to slice the project name.
      - Cleans underscores.
    - Sets:
      - `bu`, `projectName`, `extCode`.
      - `buDescription` from `BU_DESCRIPTION_MAP`.
      - `projectDescription` from `EXT_DESCRIPTION_MAP`.
  - Returns a list of `ServiceTeam` entities with descriptive fields filled.

### 6.3 `ExcelWriter` (ExtCode)

- Location: `src/main/java/invoicing/service/ext/ExcelWriter.java`
- Responsibilities:
  - Generate an ExtCode forecast Excel file summarizing:
    - Project client.
    - EXT code.
    - Description.
    - EUR amounts.
    - BU information.
  - Compute grand totals and derived financial metrics (COGS, G&A, TP).

Key logic:

- `write(List<ServiceTeam> items, File targetFolder)`
  - Creates a workbook with one sheet named `<Month> yyyy`.
  - Writes header row (Project Client, Project EXT, Descr EXT, Total EUR, BU).
  - For each `ServiceTeam`:
    - Writes project name, ext code, description, cost (as numeric), BU description.
  - Computes `grandTotal` as sum of all costs.
  - Writes:
    - “Grand Total” row.
    - COGS row (`grandTotal`).
    - G&A row (`10% of COGS`).
    - TP row (`5.5% of COGS + G&A`).
    - Total Cost row (`COGS + G&A + TP`).
  - Auto‑sizes columns.
  - Saves as `ForeCast IT <Month>.xlsx` inside `targetFolder`.

### 6.4 `FileChooserService`

- Location: `src/main/java/invoicing/service/ext/FileChooserService.java`
- Responsibilities:
  - Provide Swing‑based dialogs for selecting input files and output directory.
  - Not used by the new `InvoicingDashboard` (which implements its own choosers), but can be reused.

---

## 7. Month Module (`invoicing.service.month`)

This module splits a large “Facturación” Excel file into per‑service‑team workbooks for several months.

### 7.1 Interfaces (abstractions)

- `DateProvider`
  - Provides date‑related information:
    - `getCurrentDate(path)` extracts a date from the input file name.
    - Methods to get month names in English and Spanish.
    - `getCurrentDateTime()` for timestamping output directories.

- `ExcelFileNameGenerator`
  - Provides:
    - sheet name constants: `SHEET_AJUSTES`, `SHEET_SERVICE_HOURS_DETAILS`, `SHEET_HORAS_SERVICIO`, `SHEET_FACTURACIÓN`.
    - `generateOutputFileName(month, year, serviceTeam, directory)` to build output filenames.

- `ExcelReader`
  - `getSheet(Workbook workbook, String sheetName)` abstraction to find a sheet.

- `ExcelWriter`
  - `createWorkbookWithSheets(List<String> monthNames)`.
  - `getTotalServiceTeam(...)`.
  - `copyServiceHoursSheetData(...)` to copy and transform data from the input workbook into the output workbook.

- `ServiceTeamExtractor`
  - `extractFullServiceTeamNames(Sheet, Workbook)` returns full raw labels for service teams.
  - `extractServiceTeamNames(List<String>)` returns simplified service team names.

- `RateTable`
  - Maps approximate rates to categories and exact rate values.
  - Exposes:
    - `getCategory(double approximateRate)`.
    - `exactRate(double approximateRate)`.

- `SheetNames`
  - Value object storing names of related sheets:
    - `Horas servicio` current/next/nextNext.
    - `Ajustes`.
    - `Facturación` current/next/nextNext.
    - `Service Hours Details` current/next/nextNext.

### 7.2 Implementations (`invoicing.service.month.impl`)

- `DefaultDateProvider`
  - Implements `DateProvider`.
  - `getCurrentDate(String path)`:
    - Parses `year` and `month` from the file name (split by “-”).
  - Provides month names in English and Spanish using `Locale`.
  - Provides formatted timestamps for versioned output directories.

- `DefaultExcelFileNameGenerator`
  - Implements `ExcelFileNameGenerator`.
  - Builds filenames like: `<directory>/<month>_<year>_<serviceTeam>.xlsx`.

- `DefaultExcelReader`
  - Implements `ExcelReader`.
  - Obtains a sheet by name or throws if not found.

- `ServiceTeamExtractorImpl`
  - Implements `ServiceTeamExtractor`.
  - `extractFullServiceTeamNames(...)`:
    - Scans a sheet and finds rows whose cell B has white font color.
    - Those rows represent service team headers.
  - `extractServiceTeamNames(...)`:
    - Splits header strings by right angle bracket and returns the second part as service team name.

- `DefaultExcelWriter`
  - Implements `ExcelWriter`.
  - Uses `CogsHelper`, `Helper`, `CogsRecord`, `FiscalYear` and many Apache POI utilities.
  - Responsibilities:
    - Create an output workbook with “Service Hours Details <Month>” sheets.
    - For each service team:
      - Copy hours and absence information from “Horas servicio” and other sheets.
      - Decorate cells with styles:
        - Header, weekend, vacation, legal absence, sick leave, freeday, date style, footer currency style.
      - Add computed columns:
        - Working hours.
        - Cost per row (formula combining rate and hours).
      - Append adjustments from `Ajustes` sheet, with hourly rate and working hours.
      - Add a final “Total” row with:
        - Sum of hours.
        - Total cost based on `Facturación` sheet.
    - Helper methods:
      - `getAllData(...)`, `filterRowsByServiceTeam(...)`, `transformRows(...)`:
        - Group and transform raw rows into consolidated rows per employee.
      - `getTotalServiceTeam(...)`:
        - Compute total cost for a service team based on the `Facturación` sheet.
      - `getExactValueFromSheet(...)`:
        - Retrieve an exact numeric value for a row with a given description.

### 7.3 `ExecuteService`

- Location: `src/main/java/invoicing/service/month/ExecuteService.java`
- Responsibilities:
  - High‑level controller for the Month module.
  - Given:
    - Input Excel file path.
    - Output folder path.
    - Number of months to process.
  - Steps:
    1. Validate the input file (exists, is a file, has Excel extension).
    2. Open the workbook with Apache POI.
    3. Use `DateProvider` to compute:
       - Current date from the file name.
       - Year, numeric month, Spanish month name.
    4. Create an output directory under the given output folder:
       - `outputGeneratedExcelsFilePath + "/Version_" + currentDateTime`.
    5. Read the `Facturación <month_spanish>` sheet.
    6. Use `ServiceTeamExtractor` to:
       - Get full service team names.
       - Get simplified service team names.
    7. For each service team:
       - Build a list of month names (English) to process.
       - Use `ExcelWriter.createWorkbookWithSheets(monthNames)` to create an output workbook.
       - For each month:
         - Use `ExcelWriter.copyServiceHoursSheetData(...)` to copy and transform data from the input workbook.
       - Use `ExcelFileNameGenerator.generateOutputFileName(...)` to compute output path.
       - Use `Helper.writeWorkbook(...)` to save the workbook.

- Static convenience methods:
  - `executeScript(String inputExcelFilePath, String outputExcelsFilePath, int monthsToProcess)`
    - Creates an `ExecuteService` with default implementations.
    - Calls `process(...)`.
  - `executeScript(String inputExcelFilePath, String outputExcelsFilePath)`
    - Overload defaulting to 3 months.

---

## 8. Helper Package (`invoicing.Helper`)

### 8.1 `ReferenceData`

- Reads `Data.xlsx` (reference rate table) from disk.
- Builds maps:
  - `groupToRate`: GroupId → BigDecimal rate.
  - `rateToGroup`: rate → GroupId.
- Provides:
  - `getRateByGroup(String group)`.
  - `getGroupByApproximateRate(BigDecimal rate)`:
    - Finds closest matching rate within a tolerance.
  - `getCorrectRateByApproximate(BigDecimal rate)`.
- Used primarily by:
  - `InputRowProcessor` (Rate module).
  - `UnifiedExecutionService` → Rate module.

### 8.2 `GroupAggregator`

- Aggregates hours and cost by GroupId and by user.
- Internal maps:
  - `groupToUserHoras` and `groupToUserCost`.
- API:
  - `add(group, user, horas, cost)`.
  - `getAggregates()`: returns deep copy of hours map.
  - `getCostAggregates()`: returns deep copy of cost map.
- Used by:
  - `InputFilesReader` to accumulate values.
  - `OutputWriter` to build consolidated report.
  - `UnifiedExecutionService` for Rate module.

### 8.3 `ExcelStyler`

- Applies styling to an existing workbook:
  - Header style (bold, white text, black background).
  - Body style (bordered cells).
  - Converts column D strings to numbers.
  - Auto‑sizes columns.
- Used when styling output Excel files where needed.

### 8.4 `ServiceTeamMaps`

- Stores static maps:
  - `BU_DESCRIPTION_MAP`: BU code → BU description.
  - `EXT_DESCRIPTION_MAP`: EXT code → project description.
- Used by:
  - `ServiceTeamParser` to enrich `ServiceTeam` entities with description fields.

### 8.5 `CogsHelper`

- Works with COGS reference data loaded from `Data.xlsx` on the classpath.
- Main responsibilities:
  - `loadFromResources()`:
    - Loads COGS records from `Data.xlsx` in `src/main/resources`.
    - Builds `CogsRecord` list with GroupId, FY25, FY26 values.
  - `findGroupIdsByRate(rate, fiscalYear, records)`:
    - Returns list of GroupIds whose rate matches the whole number part of the input.
- Used by:
  - `DefaultExcelWriter` in the Month module for mapping service rates to groups.

### 8.6 `Helper`

- Collection of general helper functions:
  - `isRowEmpty(Row row)`: check if a row has any non‑blank cells.
  - `writeWorkbook(Workbook workbook, String fileName)`:
    - Writes workbook to disk.
    - Retries on “file is being used” errors.
    - Falls back to a timestamped temporary filename when needed.
  - `getDesktopPath()`: returns the user’s Desktop folder path (best effort).
  - Multiple `getXStyle` methods:
    - `getCenterStandardStyle`, `getRightStandardStyle`, `getLeftStandardStyle`.
    - `getCurrencyStyle`, `getWeekendStyle`, `getHeaderStyle`, `getLegalAbsenceStyle`, `getSickLeaveStyle`, `getFreedayStyle`, `getVacanceStyle`, `getDateStyle`, `getFooterCurrencyStyle`.
    - Used to format headers, weekends, absences, dates and currency.
  - `numberOfDays(String sheetName)`:
    - Derives month name from the sheet name and returns the number of days in that month.
  - `getRates(String input)`:
    - Parses a complex text string to extract a numeric rate in EUR.
  - `getColumnLetter(int columnIndex)`:
    - Converts zero‑based column index to Excel column letters.
  - `getMonthFromSheetName(String invoicingSheetName)`:
    - Extracts the month number from an invoicing sheet name.
- Widely used in the Month module for Excel manipulation.

---

## 9. Domain Entities and Enums

### 9.1 `invoicing.entities.ServiceTeam`

- Represents a service team entry with:
  - `bu`, `projectName`, `extCode`, `cost`, `style`.
  - `projectDescription`, `buDescription`.
- Used by:
  - `ServiceTeamParser` (to populate fields).
  - `service.ext.ExcelWriter` (to generate output Excel).

### 9.2 `invoicing.entities.CogsRecord`

- Represents a row in the COGS reference data:
  - `groupId`.
  - `fy25` and `fy26` as `BigDecimal`.
- Used by:
  - `CogsHelper` and then `DefaultExcelWriter`.

### 9.3 `invoicing.enums.FiscalYear`

- Enum with values `FY25`, `FY26`.
- Used to select which fiscal year column to read from `CogsRecord`.

---

## 10. Resources

- `Data.xlsx`
  - Reference table for rates by GroupId and COGS.
  - Used by:
    - `ReferenceData` (Rate module).
    - `CogsHelper` (Month module).

- `history.csv`
  - Stores the last used output directory and months setting.
  - Used by:
    - `InvoicingDashboard.AllInOnePanel` (append and load).
    - Legacy `UnifiedMain.AllInOnePanel`.

- `README`
  - Original minimal notes about steps (“Clean Code”, “Create executable”).

---

## 11. Overall Data Flow Summary

1. User selects input files and target directory in `InvoicingDashboard`.
2. `UnifiedExecutionService.runUnified(...)` is called.
3. Rate module:
   - Reads invoices.
   - Uses `ReferenceData` to map rates to GroupIds.
   - Aggregates with `GroupAggregator`.
   - Writes consolidated rate forecast with `OutputWriter`.
4. ExtCode module:
   - Reads service team labels and costs with `service.ext.ExcelReader`.
   - Parses them into `ServiceTeam` entities with `ServiceTeamParser` and `ServiceTeamMaps`.
   - Writes ExtCode forecast with `service.ext.ExcelWriter`.
5. Month module:
   - Uses `ExecuteService`, `DefaultDateProvider`, `DefaultExcelReader`, `ServiceTeamExtractorImpl`, `DefaultExcelWriter`.
   - Splits the “Facturación” workbook into per‑service‑team workbooks for multiple months.
   - Writes them into the month output folder.
6. All steps report logs and progress back through `UnifiedExecutionService.Listener`, which the UI uses to update the progress bar and log area.

