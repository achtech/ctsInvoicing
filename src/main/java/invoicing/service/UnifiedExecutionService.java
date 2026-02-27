package invoicing.service;

import invoicing.entities.ServiceTeam;
import invoicing.Helper.GroupAggregator;
import invoicing.Helper.ReferenceData;
import invoicing.service.ext.ExcelReader;
import invoicing.service.ext.ServiceTeamParser;
import invoicing.service.month.ExecuteService;
import invoicing.service.rate.InputFilesReader;
import invoicing.service.rate.InputRowProcessor;
import invoicing.service.rate.OutputWriter;

import java.io.File;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class UnifiedExecutionService {

    public interface Listener {
        void log(String message);
        void setProgress(int value, String barLabel, String detail);
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

        listener.setProgress(3, "Completed", "All modules finished successfully.");
        listener.log("\n=== EXECUTION COMPLETED ===");

        return mainOutputFolder;
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

