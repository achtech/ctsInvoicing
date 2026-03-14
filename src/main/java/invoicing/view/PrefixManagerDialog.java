package invoicing.view;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.prefs.Preferences;

public class PrefixManagerDialog extends JDialog {

    public static final String PREFS_KEY = "codePrefixes";
    public static final List<String> DEFAULT_PREFIXES = List.of("EXT", "INS", "INT");

    private final Preferences prefs;
    private final DefaultListModel<String> listModel = new DefaultListModel<>();
    private final JList<String> prefixList = new JList<>(listModel);
    private final JTextField inputField = new JTextField(10);

    public PrefixManagerDialog(Frame owner) {
        super(owner, "Manage Code Prefixes", true);
        this.prefs = Preferences.userNodeForPackage(InvoicingDashboard.class);

        setSize(360, 380);
        setLocationRelativeTo(owner);
        setLayout(new BorderLayout(10, 10));
        getRootPane().setBorder(new EmptyBorder(14, 14, 14, 14));

        // load saved
        loadIntoModel();

        // list
        prefixList.setFont(new Font("Monospaced", Font.PLAIN, 13));
        JScrollPane scroll = new JScrollPane(prefixList);
        scroll.setBorder(BorderFactory.createTitledBorder("Active prefixes"));
        add(scroll, BorderLayout.CENTER);

        // bottom: add + remove
        JPanel bottom = new JPanel(new BorderLayout(8, 6));
        bottom.setOpaque(false);

        JPanel addRow = new JPanel(new BorderLayout(6, 0));
        addRow.setOpaque(false);
        inputField.setToolTipText("Enter a new prefix (e.g. MSP)");
        JButton addBtn = new JButton("Add");
        addBtn.addActionListener(e -> addPrefix());
        inputField.addActionListener(e -> addPrefix()); // Enter key
        addRow.add(new JLabel("New prefix:"), BorderLayout.WEST);
        addRow.add(inputField, BorderLayout.CENTER);
        addRow.add(addBtn, BorderLayout.EAST);

        JButton removeBtn = new JButton("Remove selected");
        removeBtn.addActionListener(e -> removeSelected());

        JButton resetBtn = new JButton("Reset to defaults");
        resetBtn.addActionListener(e -> resetDefaults());

        JPanel btnRow = new JPanel(new FlowLayout(FlowLayout.RIGHT, 6, 0));
        btnRow.setOpaque(false);
        btnRow.add(resetBtn);
        btnRow.add(removeBtn);

        bottom.add(addRow, BorderLayout.NORTH);
        bottom.add(btnRow, BorderLayout.SOUTH);
        add(bottom, BorderLayout.SOUTH);
    }

    private void loadIntoModel() {
        listModel.clear();
        for (String p : getSavedPrefixes(prefs)) {
            listModel.addElement(p);
        }
    }

    private void addPrefix() {
        String val = inputField.getText().trim().toUpperCase();
        if (val.isEmpty()) return;
        if (listModel.contains(val)) {
            JOptionPane.showMessageDialog(this, "Prefix already exists: " + val);
            return;
        }
        listModel.addElement(val);
        inputField.setText("");
        save();
    }

    private void removeSelected() {
        List<String> selected = prefixList.getSelectedValuesList();
        for (String s : selected) {
            if (DEFAULT_PREFIXES.contains(s)) {
                JOptionPane.showMessageDialog(this, "Cannot remove default prefix: " + s);
                continue;
            }
            listModel.removeElement(s);
        }
        save();
    }

    private void resetDefaults() {
        listModel.clear();
        DEFAULT_PREFIXES.forEach(listModel::addElement);
        save();
    }

    private void save() {
        List<String> all = new ArrayList<>();
        for (int i = 0; i < listModel.size(); i++) all.add(listModel.getElementAt(i));
        prefs.put(PREFS_KEY, String.join(",", all));
        try { prefs.flush(); } catch (Exception ignored) {}
    }

    // ── static util used by ServiceTeamParser ──────────────────────────────
    public static String[] getSavedPrefixes(Preferences prefs) {
        String stored = prefs.get(PREFS_KEY, "");
        if (stored.isBlank()) return DEFAULT_PREFIXES.toArray(new String[0]);
        return Arrays.stream(stored.split(","))
                     .map(String::trim)
                     .filter(s -> !s.isEmpty())
                     .toArray(String[]::new);
    }
}