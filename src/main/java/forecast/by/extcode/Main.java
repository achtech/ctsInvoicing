package forecast.by.extcode;

import forecast.by.extcode.service.ExcelReader;
import forecast.by.extcode.service.ExcelWriter;
import forecast.by.extcode.service.FileChooserService;
import forecast.by.extcode.service.ServiceTeamParser;
import forecast.by.extcode.util.ServiceTeam;

import java.io.File;
import java.util.*;
import javax.swing.*;

public class Main {

    public static void main(String[] args) throws Exception {

        FileChooserService chooser = new FileChooserService();
        ExcelReader reader = new ExcelReader();
        ServiceTeamParser parser = new ServiceTeamParser();
        ExcelWriter writer = new ExcelWriter();

        List<File> files = chooser.chooseExcelFiles();
        List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();

        // Load Excel files
        for (File f : files) {
            rawItems.addAll(reader.extractRawServiceTeams(f));
        }

        // Extract labels
        List<String> labels = new ArrayList<>();
        for (ExcelReader.ServiceTeamRaw raw : rawItems) {
            labels.add(raw.getLabel());
        }

        // Parse service teams
        List<ServiceTeam> parsed = parser.parse(labels);

        // Attach cost + style info
        for (int i = 0; i < parsed.size(); i++) {

            parsed.get(i).setCost(
                    rawItems.get(i).getCost() == null
                            ? ""
                            : String.valueOf(rawItems.get(i).getCost())
            );

            parsed.get(i).setStyle(
                    rawItems.get(i).getCost() == null
                            ? null
                            : rawItems.get(i).getStyle()
            );
        }

        // Choose output folder
        File target = chooser.chooseTargetDirectory();
        if (target != null) {
            writer.write(parsed, target);   // styling applied here
            JOptionPane.showMessageDialog(null, "Excel exported successfully!");
        }
    }
}
