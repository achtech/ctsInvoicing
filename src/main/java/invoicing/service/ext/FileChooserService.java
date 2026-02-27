package invoicing.service.ext;

import javax.swing.*;
import java.io.File;
import java.util.*;

public class FileChooserService {

    public List<File> chooseExcelFiles() {
        JFileChooser chooser = new JFileChooser();
        chooser.setMultiSelectionEnabled(true);
        int result = chooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            return Arrays.asList(chooser.getSelectedFiles());
        }
        return Collections.emptyList();
    }

    public File chooseTargetDirectory() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = chooser.showSaveDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            return chooser.getSelectedFile();
        }
        return null;
    }
}
