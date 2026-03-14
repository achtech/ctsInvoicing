// package invoicing.service.ext;


// import java.util.*;

// import invoicing.entities.ServiceTeam;
// import invoicing.Helper.ServiceTeamMaps;

// public class ServiceTeamParser {

//     public List<ServiceTeam> parse(List<String> rawItems) {
//         List<ServiceTeam> list = new ArrayList<>();

//         for (String b : rawItems) {
//             ServiceTeam st = new ServiceTeam();
//             String t[] = b.split(" > ");
//             String s = t.length>1 ? t[1] : t.length == 1 ? t[0] : "";
//             String bu = extractBU(s);
//             String ext = extractEXT(s);
//             String projectName = extractProjectName(s, bu, ext);

//             st.setBu(bu);
//             st.setProjectName(projectName);
//             st.setExtCode(ext);
//             st.setBuDescription(ServiceTeamMaps.BU_DESCRIPTION_MAP.getOrDefault(bu, ""));
//             st.setProjectDescription(ServiceTeamMaps.EXT_DESCRIPTION_MAP.getOrDefault(ext, ""));
//             if(!projectName.isEmpty()) list.add(st);
//         }
//         return list;
//     }

//     private String extractBU(String s) {
//         for (String key : ServiceTeamMaps.BU_DESCRIPTION_MAP.keySet()) {
//             if (s.startsWith(key)) return key;
//         }
//         return "";
//     }

//     private String extractEXT(String s) {
//         int idx = s.indexOf("EXT");
//         if (idx != -1) return s.substring(idx).trim();
//         return "Pending";
//     }

//     private String extractProjectName(String s, String bu, String ext) {
//         int start = bu.isEmpty() ? 0 : s.indexOf(bu) + bu.length() + 1;
//         int end = ext.equals("Pending") ? s.length() : s.indexOf(ext) - 1;
//         if (start < 0 || end < 0 || start >= end) return "";
//         String str = s.substring(start, end).trim();
//         str = str.replaceAll("^_+|_+$", "").replace('_', ' ');
//         return str;
//     }
// }
package invoicing.service.ext;

import invoicing.Helper.ServiceTeamMaps;
import invoicing.entities.ServiceTeam;
import invoicing.view.PrefixManagerDialog;

import java.util.*;
import java.util.prefs.Preferences;

public class ServiceTeamParser {

    private final String[] CODE_PREFIXES;

    public ServiceTeamParser() {
        Preferences prefs = Preferences.userNodeForPackage(invoicing.view.InvoicingDashboard.class);
        CODE_PREFIXES = PrefixManagerDialog.getSavedPrefixes(prefs);
    }

    public List<ServiceTeam> parse(List<String> rawItems) {
        List<ServiceTeam> list = new ArrayList<>();

        for (String b : rawItems) {
            ServiceTeam st = new ServiceTeam();
            String[] t = b.split(" > ");
            String s = t.length > 1 ? t[1] : t.length == 1 ? t[0] : "";
            if (s.isBlank()) continue;

            String bu   = extractBU(s);
            String code = extractCode(s);
            if (code.equals("Pending") && t.length > 2) code = extractCode(t[2]);
            if (code.equals("Pending"))                  code = extractCode(b);

            String projectName = extractProjectName(s, bu, code);

            st.setBu(bu);
            st.setProjectName(projectName);
            st.setExtCode(code);
            st.setBuDescription(ServiceTeamMaps.BU_DESCRIPTION_MAP.getOrDefault(bu, ""));
            st.setProjectDescription(ServiceTeamMaps.CODE_DESCRIPTION_MAP.getOrDefault(code, ""));
            if (!projectName.isEmpty()) list.add(st);
        }
        return list;
    }

    private String extractBU(String s) {
        for (String key : ServiceTeamMaps.BU_DESCRIPTION_MAP.keySet()) {
            if (s.startsWith(key)) return key;
        }
        return "";
    }

    private String extractCode(String s) {
    for (String prefix : CODE_PREFIXES) {
        int idx = s.indexOf(prefix);
        if (idx != -1) {
            String raw = s.substring(idx).split("\\s+")[0].trim();
            raw = raw.replace('_', '-');
            int parenIdx = raw.indexOf('(');
            if (parenIdx != -1) raw = raw.substring(0, parenIdx);
            return raw;
        }
    }
    return "Pending";
}

    private String extractProjectName(String s, String bu, String code) {
        int start = bu.isEmpty() ? 0 : s.indexOf(bu) + bu.length() + 1;
        int codeStart = -1;
        for (String prefix : CODE_PREFIXES) {
            int idx = s.indexOf(prefix);
            if (idx != -1) { codeStart = idx; break; }
        }
        int end = codeStart == -1 ? s.length() : codeStart - 1;
        if (start < 0 || end < 0 || start >= end) return "";
        String str = s.substring(start, end).trim();
        str = str.replaceAll("^[_\\-]+|[_\\-]+$", "").replace('_', ' ');
        return str;
    }
}