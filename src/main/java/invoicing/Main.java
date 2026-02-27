package invoicing;

import invoicing.view.InvoicingDashboard;

import javax.swing.*;
import java.awt.*;

public class Main {

    private static final Font MAIN_FONT = new Font("Segoe UI", Font.PLAIN, 13);

    public static void main(String[] args) {
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (Exception e) {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

        UIManager.put("Label.font", MAIN_FONT);
        UIManager.put("Button.font", MAIN_FONT);
        UIManager.put("TextField.font", MAIN_FONT);
        UIManager.put("TextArea.font", new Font("Monospaced", Font.PLAIN, 12));

        SwingUtilities.invokeLater(() -> new InvoicingDashboard().setVisible(true));
    }
}

