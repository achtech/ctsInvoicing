package invoicing.service;

import invoicing.entities.ServiceTeam;
import invoicing.Helper.GroupAggregator;
import invoicing.Helper.ReferenceData;
import invoicing.service.global.GlobalService;
import invoicing.service.ext.ExcelReader;
import invoicing.service.ext.ServiceTeamParser;
import invoicing.service.month.ExecuteService;
import invoicing.service.rate.InputFilesReader;
import invoicing.service.rate.InputRowProcessor;
import invoicing.service.rate.OutputWriter;

import java.io.File;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class UnifiedExecutionService {

    public interface Listener {
        void log(String message);
        void setProgress(int value, String barLabel, String detail);
    }

    public static final class RateRow {
        private final String groupId;
        private final BigDecimal rate;
        private final double hours;
        private final double cost;

        public RateRow(String groupId, BigDecimal rate, double hours, double cost) {
            this.groupId = groupId;
            this.rate = rate;
            this.hours = hours;
            this.cost = cost;
        }

        public String getGroupId() {
            return groupId;
        }

        public BigDecimal getRate() {
            return rate;
        }

        public double getHours() {
            return hours;
        }

        public double getCost() {
            return cost;
        }
    }

    public static final class ExtRow {
        private final String projectName;
        private final String extCode;
        private final String projectDescription;
        private final String buDescription;
        private final double cost;

        public ExtRow(String projectName, String extCode, String projectDescription, String buDescription, double cost) {
            this.projectName = projectName;
            this.extCode = extCode;
            this.projectDescription = projectDescription;
            this.buDescription = buDescription;
            this.cost = cost;
        }

        public String getProjectName() {
            return projectName;
        }

        public String getExtCode() {
            return extCode;
        }

        public String getProjectDescription() {
            return projectDescription;
        }

        public String getBuDescription() {
            return buDescription;
        }

        public double getCost() {
            return cost;
        }
    }

    public static final class MonthInputData {
        private final String inputFileName;
        private final int monthsUsed;
        private final List<ExecuteService.MonthWorkbookData> workbooks;
        private final String error;

        public MonthInputData(String inputFileName, int monthsUsed, List<ExecuteService.MonthWorkbookData> workbooks) {
            this(inputFileName, monthsUsed, workbooks, null);
        }

        public MonthInputData(String inputFileName, int monthsUsed, List<ExecuteService.MonthWorkbookData> workbooks, String error) {
            this.inputFileName = inputFileName;
            this.monthsUsed = monthsUsed;
            this.workbooks = workbooks;
            this.error = error;
        }

        public String getInputFileName() {
            return inputFileName;
        }

        public int getMonthsUsed() {
            return monthsUsed;
        }

        public List<ExecuteService.MonthWorkbookData> getWorkbooks() {
            return workbooks;
        }

        public String getError() {
            return error;
        }

        public boolean isSuccess() {
            return error == null || error.isBlank();
        }

        public int getWorkbookCount() {
            return workbooks == null ? 0 : workbooks.size();
        }

        @Override
        public String toString() {
            return "MonthInputData{" +
                    "inputFileName='" + inputFileName + '\'' +
                    ", monthsUsed=" + monthsUsed +
                    ", workbooks=" + getWorkbookCount() +
                    (isSuccess() ? "" : ", error='" + error + '\'') +
                    '}';
        }
    }

    public static final class MonthServiceTeamSheet {
        private final String serviceTeam;
        private final ExecuteService.SheetTableData table;

        public MonthServiceTeamSheet(String serviceTeam, ExecuteService.SheetTableData table) {
            this.serviceTeam = serviceTeam;
            this.table = table;
        }

        public String getServiceTeam() {
            return serviceTeam;
        }

        public ExecuteService.SheetTableData getTable() {
            return table;
        }
    }

    public static final class ProjectMonthGroup {
        private final String projectName;
        private final int monthsUsed;
        private final String error;
        private final List<MonthGroup> months;

        public ProjectMonthGroup(String projectName, int monthsUsed, String error, List<MonthGroup> months) {
            this.projectName = projectName;
            this.monthsUsed = monthsUsed;
            this.error = error;
            this.months = months;
        }

        public String getProjectName() {
            return projectName;
        }

        public int getMonthsUsed() {
            return monthsUsed;
        }

        public String getError() {
            return error;
        }

        public List<MonthGroup> getMonths() {
            return months;
        }
    }

    public static final class MonthGroup {
        private final String month;
        private final List<MonthServiceTeamSheet> serviceTeams;

        public MonthGroup(String month, List<MonthServiceTeamSheet> serviceTeams) {
            this.month = month;
            this.serviceTeams = serviceTeams;
        }

        public String getMonth() {
            return month;
        }

        public List<MonthServiceTeamSheet> getServiceTeams() {
            return serviceTeams;
        }
    }

    public static final class GlobalMonthGroup {
        private final String month;
        private final List<GlobalProjectGroup> projects;

        public GlobalMonthGroup(String month, List<GlobalProjectGroup> projects) {
            this.month = month;
            this.projects = projects;
        }

        public String getMonth() {
            return month;
        }

        public List<GlobalProjectGroup> getProjects() {
            return projects;
        }
    }

    public static final class GlobalProjectGroup {
        private final String projectName;
        private final List<MonthServiceTeamSheet> serviceTeams;

        public GlobalProjectGroup(String projectName, List<MonthServiceTeamSheet> serviceTeams) {
            this.projectName = projectName;
            this.serviceTeams = serviceTeams;
        }

        public String getProjectName() {
            return projectName;
        }

        public List<MonthServiceTeamSheet> getServiceTeams() {
            return serviceTeams;
        }
    }

    public static final class UnifiedDataResult {
        private final List<RateRow> rateRows;
        private final List<ExtRow> extRows;
        private final List<MonthInputData> monthInputs;
        private final List<ExecuteService.SheetTableData> globalMonthSheets;
        private final List<ProjectMonthGroup> monthByProject;
        private final List<GlobalMonthGroup> globalByMonth;

        public UnifiedDataResult(
                List<RateRow> rateRows,
                List<ExtRow> extRows,
                List<MonthInputData> monthInputs,
                List<ExecuteService.SheetTableData> globalMonthSheets,
                List<ProjectMonthGroup> monthByProject,
                List<GlobalMonthGroup> globalByMonth
        ) {
            this.rateRows = rateRows;
            this.extRows = extRows;
            this.monthInputs = monthInputs;
            this.globalMonthSheets = globalMonthSheets;
            this.monthByProject = monthByProject;
            this.globalByMonth = globalByMonth;
        }

        public List<RateRow> getRateRows() {
            return rateRows;
        }

        public List<ExtRow> getExtRows() {
            return extRows;
        }

        public List<MonthInputData> getMonthInputs() {
            return monthInputs;
        }

        public List<ExecuteService.SheetTableData> getGlobalMonthSheets() {
            return globalMonthSheets;
        }

        public List<ProjectMonthGroup> getMonthByProject() {
            return monthByProject;
        }

        public List<GlobalMonthGroup> getGlobalByMonth() {
            return globalByMonth;
        }
    }

    public File runUnified(File targetDir, List<File> inputs, int months, boolean useManual, Listener listener) {
        LocalDateTime now = LocalDateTime.now();
        String currentMonthStr = now.format(DateTimeFormatter.ofPattern("MMM_yyyy"));
        String runStamp = now.format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        File mainOutputFolder = new File(targetDir, "forecast_italy_" + currentMonthStr + "_" + runStamp);
        mainOutputFolder.mkdirs();

        File rateFolder = new File(mainOutputFolder, "forecast_it_rate_" + currentMonthStr);
        File extFolder = new File(mainOutputFolder, "forecast_EXT_" + currentMonthStr);
        File monthFolder = new File(mainOutputFolder, "forecast_month_" + currentMonthStr);
        rateFolder.mkdirs();
        extFolder.mkdirs();
        monthFolder.mkdirs();

        listener.log("=== STARTING UNIFIED EXECUTION ===");
        listener.log("Output Folder : " + mainOutputFolder.getAbsolutePath());
        listener.log("Months mode   : " + (useManual ? "MANUAL (" + months + " months)" : "AUTO-DETECT from Facturacion sheets"));

        runRateModule(now, rateFolder, inputs, listener);
        runExtModule(extFolder, inputs, listener);
        runMonthModule(monthFolder, inputs, months, useManual, listener);
        runGlobalMonthConsolidation(monthFolder, now, listener);

        listener.setProgress(3, "Completed", "All modules finished successfully.");
        listener.log("\n=== EXECUTION COMPLETED ===");

        return mainOutputFolder;
    }

    public UnifiedDataResult runUnifiedData(List<File> inputs, int months, boolean useManual, Listener listener) {
        listener.log("=== STARTING UNIFIED DATA MODE ===");
        listener.log("Months mode   : " + (useManual ? "MANUAL (" + months + " months)" : "AUTO-DETECT from Facturacion sheets"));

        listener.setProgress(0, "Step 1/3 — Rate", "Computing Rate data...");
        List<RateRow> rateRows = computeRateData(inputs, listener);

        listener.setProgress(1, "Step 2/3 — ExtCode", "Computing ExtCode data...");
        List<ExtRow> extRows = computeExtData(inputs, listener);

        listener.setProgress(2, "Step 3/3 — Month", "Computing Month data...");
        List<MonthInputData> monthInputs = computeMonthData(inputs, months, useManual, listener);
        List<ExecuteService.SheetTableData> globalMonthSheets = consolidateMonthData(monthInputs);
        List<ProjectMonthGroup> monthByProject = buildMonthByProject(monthInputs);
        List<GlobalMonthGroup> globalByMonth = buildGlobalByMonth(monthInputs);

        listener.setProgress(3, "Completed", "Data extraction completed.");
        listener.log("=== DATA MODE COMPLETED ===");

        return new UnifiedDataResult(rateRows, extRows, monthInputs, globalMonthSheets, monthByProject, globalByMonth);
    }

    private List<ProjectMonthGroup> buildMonthByProject(List<MonthInputData> monthInputs) {
        if (monthInputs == null || monthInputs.isEmpty()) return Collections.emptyList();

        List<ProjectMonthGroup> out = new ArrayList<>();
        for (MonthInputData input : monthInputs) {
            if (input == null) continue;
            if (!input.isSuccess()) {
                out.add(new ProjectMonthGroup(input.getInputFileName(), input.getMonthsUsed(), input.getError(), Collections.emptyList()));
                continue;
            }

            Map<String, List<MonthServiceTeamSheet>> byMonth = new LinkedHashMap<>();
            List<ExecuteService.MonthWorkbookData> workbooks = input.getWorkbooks();
            if (workbooks != null) {
                for (ExecuteService.MonthWorkbookData wb : workbooks) {
                    if (wb == null || wb.getSheets() == null) continue;
                    String serviceTeam = wb.getServiceTeam();
                    for (ExecuteService.SheetTableData sheet : wb.getSheets()) {
                        if (sheet == null) continue;
                        String monthKey = extractMonthKey(sheet.getSheetName());
                        byMonth.computeIfAbsent(monthKey, ignored -> new ArrayList<>())
                                .add(new MonthServiceTeamSheet(serviceTeam, sheet));
                    }
                }
            }

            List<MonthGroup> months = new ArrayList<>();
            for (Map.Entry<String, List<MonthServiceTeamSheet>> e : byMonth.entrySet()) {
                months.add(new MonthGroup(e.getKey(), e.getValue()));
            }

            out.add(new ProjectMonthGroup(input.getInputFileName(), input.getMonthsUsed(), null, months));
        }

        return out;
    }

    private List<GlobalMonthGroup> buildGlobalByMonth(List<MonthInputData> monthInputs) {
        if (monthInputs == null || monthInputs.isEmpty()) return Collections.emptyList();

        Map<String, Map<String, List<MonthServiceTeamSheet>>> byMonthByProject = new LinkedHashMap<>();

        for (MonthInputData input : monthInputs) {
            if (input == null || !input.isSuccess()) continue;
            String project = input.getInputFileName();
            List<ExecuteService.MonthWorkbookData> workbooks = input.getWorkbooks();
            if (workbooks == null) continue;

            for (ExecuteService.MonthWorkbookData wb : workbooks) {
                if (wb == null || wb.getSheets() == null) continue;
                String serviceTeam = wb.getServiceTeam();
                for (ExecuteService.SheetTableData sheet : wb.getSheets()) {
                    if (sheet == null) continue;
                    String monthKey = extractMonthKey(sheet.getSheetName());
                    byMonthByProject
                            .computeIfAbsent(monthKey, ignored -> new LinkedHashMap<>())
                            .computeIfAbsent(project, ignored -> new ArrayList<>())
                            .add(new MonthServiceTeamSheet(serviceTeam, sheet));
                }
            }
        }

        List<GlobalMonthGroup> out = new ArrayList<>();
        for (Map.Entry<String, Map<String, List<MonthServiceTeamSheet>>> monthEntry : byMonthByProject.entrySet()) {
            List<GlobalProjectGroup> projects = new ArrayList<>();
            for (Map.Entry<String, List<MonthServiceTeamSheet>> projectEntry : monthEntry.getValue().entrySet()) {
                projects.add(new GlobalProjectGroup(projectEntry.getKey(), projectEntry.getValue()));
            }
            out.add(new GlobalMonthGroup(monthEntry.getKey(), projects));
        }
        return out;
    }

    private String extractMonthKey(String sheetName) {
        if (sheetName == null) return "";
        String prefix = "Service Hours Details ";
        if (sheetName.startsWith(prefix)) return sheetName.substring(prefix.length()).trim();
        return sheetName.trim();
    }

    public UnifiedDataResult runUnifiedDataFromArgs(String[] args, int months, boolean useManual, Listener listener) {
        if (args == null || args.length == 0) {
            throw new IllegalArgumentException("No input files provided (args is empty).");
        }

        List<File> inputs = new ArrayList<>();
        for (String a : args) {
            if (a == null) continue;
            String path = a.trim();
            if (path.isEmpty()) continue;
            File f = new File(path);
            if (!f.exists() || !f.isFile()) {
                throw new IllegalArgumentException("Invalid input file: " + f.getAbsolutePath());
            }
            inputs.add(f);
        }

        if (inputs.isEmpty()) {
            throw new IllegalArgumentException("No valid input files found in args.");
        }

        return runUnifiedData(inputs, months, useManual, listener);
    }

    public UnifiedDataResult runUnifiedDataFromPaths(List<String> filePaths, int months, boolean useManual, Listener listener) {
        if (filePaths == null || filePaths.isEmpty()) {
            throw new IllegalArgumentException("No input files provided (filePaths is empty).");
        }

        List<File> inputs = new ArrayList<>();
        for (String p : filePaths) {
            if (p == null) continue;
            String path = p.trim();
            if (path.isEmpty()) continue;
            File f = new File(path);
            if (!f.exists() || !f.isFile()) {
                throw new IllegalArgumentException("Invalid input file: " + f.getAbsolutePath());
            }
            inputs.add(f);
        }

        if (inputs.isEmpty()) {
            throw new IllegalArgumentException("No valid input files found in filePaths.");
        }

        return runUnifiedData(inputs, months, useManual, listener);
    }

    private List<ExecuteService.SheetTableData> consolidateMonthData(List<MonthInputData> monthInputs) {
        if (monthInputs == null || monthInputs.isEmpty()) return Collections.emptyList();

        Map<String, List<List<String>>> rowsBySheet = new LinkedHashMap<>();
        Map<String, Integer> maxColsBySheet = new LinkedHashMap<>();
        Map<String, Integer> costColBySheet = new LinkedHashMap<>();
        Map<String, Double> totalCostBySheet = new LinkedHashMap<>();

        for (MonthInputData input : monthInputs) {
            if (input == null || input.getWorkbooks() == null) continue;

            for (ExecuteService.MonthWorkbookData wb : input.getWorkbooks()) {
                if (wb == null || wb.getSheets() == null) continue;

                String projectName = (input.getInputFileName() == null ? "" : input.getInputFileName()) +
                        " - " + (wb.getServiceTeam() == null ? "" : wb.getServiceTeam());

                for (ExecuteService.SheetTableData sheet : wb.getSheets()) {
                    if (sheet == null || sheet.getRows() == null) continue;

                    String sheetName = sheet.getSheetName() == null ? "" : sheet.getSheetName();
                    List<List<String>> srcRows = sheet.getRows();
                    int srcMaxCols = maxColumns(srcRows);

                    int prevMax = maxColsBySheet.getOrDefault(sheetName, 0);
                    int newMax = Math.max(prevMax, srcMaxCols);
                    maxColsBySheet.put(sheetName, newMax);

                    List<List<String>> globalRows = rowsBySheet.computeIfAbsent(sheetName, ignored -> new ArrayList<>());

                    List<String> titleRow = new ArrayList<>(Collections.nCopies(newMax, ""));
                    if (!titleRow.isEmpty()) titleRow.set(0, projectName);
                    globalRows.add(titleRow);

                    for (List<String> r : srcRows) {
                        globalRows.add(padRow(r, newMax));
                    }
                    globalRows.add(new ArrayList<>(Collections.nCopies(newMax, "")));

                    int costCol = costColBySheet.getOrDefault(sheetName, -1);
                    if (costCol < 0) {
                        costCol = findCostColumn(srcRows);
                        if (costCol < 0) costCol = findLastNonEmptyColumn(srcRows);
                        costColBySheet.put(sheetName, costCol);
                    }

                    Double projectTotal = extractTotalCost(srcRows, costCol);
                    if (projectTotal != null) {
                        totalCostBySheet.merge(sheetName, projectTotal, Double::sum);
                    }
                }
            }
        }

        List<ExecuteService.SheetTableData> out = new ArrayList<>();
        for (Map.Entry<String, List<List<String>>> e : rowsBySheet.entrySet()) {
            String sheetName = e.getKey();
            List<List<String>> rows = e.getValue();
            int maxCols = Math.max(1, maxColsBySheet.getOrDefault(sheetName, 1));
            int costCol = costColBySheet.getOrDefault(sheetName, Math.max(0, maxCols - 1));
            if (costCol >= maxCols) costCol = maxCols - 1;

            rows.add(new ArrayList<>(Collections.nCopies(maxCols, "")));
            List<String> totalRow = new ArrayList<>(Collections.nCopies(maxCols, ""));
            totalRow.set(0, "TOTAL COST (ALL PROJECTS)");
            totalRow.set(costCol, formatNumber(totalCostBySheet.getOrDefault(sheetName, 0.0)));
            rows.add(totalRow);

            out.add(new ExecuteService.SheetTableData(sheetName, rows));
        }

        return out;
    }

    private int maxColumns(List<List<String>> rows) {
        int max = 0;
        if (rows == null) return 0;
        for (List<String> r : rows) {
            if (r != null && r.size() > max) max = r.size();
        }
        return max;
    }

    private List<String> padRow(List<String> row, int size) {
        if (size <= 0) return Collections.emptyList();
        if (row == null) return new ArrayList<>(Collections.nCopies(size, ""));
        if (row.size() == size) return new ArrayList<>(row);
        List<String> out = new ArrayList<>(Collections.nCopies(size, ""));
        for (int i = 0; i < Math.min(size, row.size()); i++) out.set(i, row.get(i));
        return out;
    }

    private int findCostColumn(List<List<String>> rows) {
        if (rows == null || rows.isEmpty()) return -1;
        List<String> header = rows.get(0);
        if (header == null) return -1;
        for (int i = 0; i < header.size(); i++) {
            String v = header.get(i);
            if (v == null) continue;
            if (v.trim().toLowerCase().startsWith("cost")) return i;
        }
        return -1;
    }

    private int findLastNonEmptyColumn(List<List<String>> rows) {
        if (rows == null) return 0;
        for (int r = rows.size() - 1; r >= 0; r--) {
            List<String> row = rows.get(r);
            if (row == null) continue;
            for (int c = row.size() - 1; c >= 0; c--) {
                String v = row.get(c);
                if (v != null && !v.trim().isEmpty()) return c;
            }
        }
        return 0;
    }

    private Double extractTotalCost(List<List<String>> rows, int costCol) {
        if (rows == null || rows.isEmpty()) return null;
        for (int r = rows.size() - 1; r >= 0; r--) {
            List<String> row = rows.get(r);
            if (row == null) continue;
            boolean empty = true;
            for (String v : row) {
                if (v != null && !v.trim().isEmpty()) {
                    empty = false;
                    break;
                }
            }
            if (empty) continue;
            if (costCol < 0 || costCol >= row.size()) return null;
            return parseDoubleLenient(row.get(costCol));
        }
        return null;
    }

    private Double parseDoubleLenient(String s) {
        if (s == null) return null;
        String v = s.trim();
        if (v.isEmpty()) return null;

        v = v.replaceAll("[^0-9,\\.\\-]", "");
        if (v.isEmpty() || "-".equals(v)) return null;

        int lastComma = v.lastIndexOf(',');
        int lastDot = v.lastIndexOf('.');
        if (lastComma >= 0 && lastDot >= 0) {
            if (lastComma > lastDot) {
                v = v.replace(".", "");
                v = v.replace(",", ".");
            } else {
                v = v.replace(",", "");
            }
        } else if (lastComma >= 0) {
            v = v.replace(",", ".");
        }

        try {
            return Double.parseDouble(v);
        } catch (Exception ignored) {
            return null;
        }
    }

    private String formatNumber(double value) {
        return String.valueOf(Math.round(value * 100.0) / 100.0);
    }

    private List<RateRow> computeRateData(List<File> inputs, Listener listener) {
        try {
            ReferenceData referenceData = new ReferenceData();
            try (InputStream dataStream = getClass().getClassLoader().getResourceAsStream("Data.xlsx")) {
                if (dataStream == null) {
                    listener.log("  ! Rate Error: Data.xlsx not found inside the JAR. Check build resources.");
                    return Collections.emptyList();
                }
                referenceData.load(dataStream);
            } catch (Exception e) {
                listener.log("  ! Rate Error loading Data.xlsx: " + e.getMessage());
                return Collections.emptyList();
            }

            GroupAggregator aggregator = new GroupAggregator();
            InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
            InputFilesReader filesReader = new InputFilesReader(rowProcessor, aggregator);

            for (File f : inputs) {
                try {
                    filesReader.processFile(f.getAbsolutePath());
                } catch (Exception e) {
                    listener.log("  - Rate Warning: Failed to process " + f.getName());
                }
            }

            Map<String, Map<String, Double>> hoursAgg = aggregator.getAggregates();
            Map<String, Map<String, Double>> costAgg = aggregator.getCostAggregates();
            if (hoursAgg.isEmpty()) return Collections.emptyList();

            List<RateRow> out = new ArrayList<>();
            for (Map.Entry<String, Map<String, Double>> groupEntry : hoursAgg.entrySet()) {
                String groupId = groupEntry.getKey();
                BigDecimal rate = referenceData.getRateByGroup(groupId);
                if (rate == null) continue;

                double hours = 0;
                for (Double h : groupEntry.getValue().values()) {
                    hours += h == null ? 0 : h;
                }

                double cost = 0;
                Map<String, Double> usersCost = costAgg.get(groupId);
                if (usersCost != null) {
                    for (Double c : usersCost.values()) {
                        cost += c == null ? 0 : c;
                    }
                }

                out.add(new RateRow(groupId, rate, hours, cost));
            }
            return out;
        } catch (Exception e) {
            listener.log("  ! Rate Data Failed: " + e.getMessage());
            return Collections.emptyList();
        }
    }

    private List<ExtRow> computeExtData(List<File> inputs, Listener listener) {
        try {
            ExcelReader reader = new ExcelReader();
            ServiceTeamParser parser = new ServiceTeamParser();

            List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
            for (File f : inputs) {
                try {
                    rawItems.addAll(reader.extractRawServiceTeams(f));
                } catch (Exception e) {
                    listener.log("  - ExtCode Warning: Failed to process " + f.getName());
                }
            }

            if (rawItems.isEmpty()) return Collections.emptyList();

            List<String> labels = new ArrayList<>();
            for (ExcelReader.ServiceTeamRaw raw : rawItems) {
                labels.add(raw.getLabel());
            }

            List<ServiceTeam> parsed = parser.parse(labels);
            List<ExtRow> out = new ArrayList<>();

            for (int i = 0; i < parsed.size(); i++) {
                ServiceTeam st = parsed.get(i);
                ExcelReader.ServiceTeamRaw raw = rawItems.get(i);
                String rawCost = raw.getCost() == null ? "" : String.valueOf(raw.getCost());
                st.setCost(rawCost);
                st.setStyle(raw.getCost() == null ? null : raw.getStyle());

                double cost = 0;
                if (raw.getCost() != null) cost = raw.getCost();

                out.add(new ExtRow(
                        st.getProjectName(),
                        st.getExtCode(),
                        st.getProjectDescription(),
                        st.getBuDescription(),
                        cost
                ));
            }

            return out;
        } catch (Exception e) {
            listener.log("  ! ExtCode Data Failed: " + e.getMessage());
            return Collections.emptyList();
        }
    }

    private List<MonthInputData> computeMonthData(List<File> inputs, int months, boolean useManual, Listener listener) {
        List<MonthInputData> out = new ArrayList<>();
        for (File f : inputs) {
            int currentMonths = months;
            try {
                if (!useManual) {
                    int detected = countMonthSheets(f, listener);
                    currentMonths = detected > 0 ? detected : months;
                }
                List<ExecuteService.MonthWorkbookData> workbooks = ExecuteService.executeToData(f.getAbsolutePath(), currentMonths);
                out.add(new MonthInputData(f.getName(), currentMonths, workbooks));
            } catch (Exception e) {
                listener.log("  - Month Data Warning: Failed to process " + f.getName() + ": " + e.getMessage());
                out.add(new MonthInputData(f.getName(), currentMonths, Collections.emptyList(), e.toString()));
            }
        }
        return out;
    }

    private void runRateModule(LocalDateTime now, File rateFolder, List<File> inputs, Listener listener) {
        listener.setProgress(0, "Step 1/3 — Rate", "Running Forecast By Rate...");
        listener.log("\n[1/3] Running Forecast By Rate...");
        try {
            ReferenceData referenceData = new ReferenceData();
            try (InputStream dataStream = getClass().getClassLoader().getResourceAsStream("Data.xlsx")) {
                if (dataStream == null) {
                    listener.log("  ! Rate Error: Data.xlsx not found inside the JAR. Check build resources.");
                    return;
                }
                referenceData.load(dataStream);
            } catch (Exception e) {
                listener.log("  ! Rate Error loading Data.xlsx: " + e.getMessage());
                return;
            }


            GroupAggregator aggregator = new GroupAggregator();
            InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
            InputFilesReader filesReader = new InputFilesReader(rowProcessor, aggregator);

            for (File f : inputs) {
                try {
                    filesReader.processFile(f.getAbsolutePath());
                } catch (Exception e) {
                    listener.log("  - Rate Warning: Failed to process " + f.getName());
                }
            }

            if (!aggregator.getAggregates().isEmpty()) {
                OutputWriter writer = new OutputWriter(referenceData, aggregator);
                String fullMonth = now.format(DateTimeFormatter.ofPattern("MMMM"));
                String rateOut = new File(rateFolder, "Rate Forecast " + fullMonth + ".xlsx").getAbsolutePath();
                writer.write(rateOut);
                listener.log("  > Rate Report created: " + rateOut);
            } else {
                listener.log("  - Rate Warning: No valid data found for Rate module.");
            }
        } catch (Exception e) {
            listener.log("  ! Rate Module Failed: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void runExtModule(File extFolder, List<File> inputs, Listener listener) {
        listener.setProgress(1, "Step 2/3 — ExtCode", "Running Forecast By ExtCode...");
        listener.log("\n[2/3] Running Forecast By ExtCode...");
        try {
            ExcelReader reader = new ExcelReader();
            ServiceTeamParser parser = new ServiceTeamParser();
            invoicing.service.ext.ExcelWriter writer = new invoicing.service.ext.ExcelWriter();

            List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
            for (File f : inputs) {
                try {
                    rawItems.addAll(reader.extractRawServiceTeams(f));
                } catch (Exception e) {
                    listener.log("  - ExtCode Warning: Failed to process " + f.getName());
                }
            }

            if (!rawItems.isEmpty()) {
                List<String> labels = new ArrayList<>();
                for (ExcelReader.ServiceTeamRaw raw : rawItems) {
                    labels.add(raw.getLabel());
                }
                List<ServiceTeam> parsed = parser.parse(labels);
                for (int i = 0; i < parsed.size(); i++) {
                    parsed.get(i).setCost(rawItems.get(i).getCost() == null ? "" : String.valueOf(rawItems.get(i).getCost()));
                    parsed.get(i).setStyle(rawItems.get(i).getCost() == null ? null : rawItems.get(i).getStyle());
                }
                writer.write(parsed, extFolder);
                listener.log("  > ExtCode Report created in: " + extFolder.getAbsolutePath());
            } else {
                listener.log("  - ExtCode Warning: No valid data found.");
            }
        } catch (Exception e) {
            listener.log("  ! ExtCode Module Failed: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void runMonthModule(File monthFolder, List<File> inputs, int months, boolean useManual, Listener listener) {
        listener.setProgress(2, "Step 3/3 — Month", "Running Forecast By Month...");
        listener.log("\n[3/3] Running Forecast By Month...");
        try {
            for (File f : inputs) {
                try {
                    int currentMonths = months;
                    if (!useManual) {
                        int detected = countMonthSheets(f, listener);
                        if (detected > 0) {
                            currentMonths = detected;
                            listener.log("  - Auto-detected months for " + f.getName() + ": " + currentMonths);
                        } else {
                            listener.log("  - Warning: No Facturacion sheets found in " + f.getName() + ". Using default: " + months);
                        }
                    }
                    listener.log("  - Processing " + f.getName() + " with " + currentMonths + " months...");
                    ExecuteService.executeScript(f.getAbsolutePath(), monthFolder.getAbsolutePath(), currentMonths);
                } catch (Exception e) {
                    listener.log("  - Month Warning: Failed to process " + f.getName() + ": " + e.getMessage());
                }
            }
            listener.log("  > Month processing finished.");
        } catch (Exception e) {
            listener.log("  ! Month Module Critical Error: " + e.getMessage());
        }
    }

    private void runGlobalMonthConsolidation(File monthFolder, LocalDateTime now, Listener listener) {
        listener.log("\n[Global] Consolidating Month Excel outputs...");
        try {
            GlobalService globalService = new GlobalService();
            File out = globalService.generateGlobalMonthWorkbook(monthFolder, now);
            if (out == null) {
                listener.log("  - Global: No month Excel files found to consolidate.");
            } else {
                listener.log("  > Global consolidated report created: " + out.getAbsolutePath());
            }
        } catch (Exception e) {
            listener.log("  ! Global Module Failed: " + e.getMessage());
        }
    }

    private int countMonthSheets(File f, Listener listener) {
        int count = 0;
        try (org.apache.poi.ss.usermodel.Workbook wb = org.apache.poi.ss.usermodel.WorkbookFactory.create(f)) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String n = wb.getSheetName(i).toLowerCase()
                        .replace("á", "a").replace("é", "e")
                        .replace("í", "i").replace("ó", "o").replace("ú", "u");
                if (n.contains("facturacion")) {
                    count++;
                }
            }
        } catch (Exception e) {
            listener.log("  - Error counting sheets in " + f.getName() + ": " + e.getMessage());
        }
        return count;
    }
}
