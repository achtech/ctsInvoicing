package forecast.by.extcode;

import forecast.by.util.ServiceTeam;
import forecast.by.util.ServiceTeamMaps;

import java.util.*;

public class ServiceTeamParser {

    public List<ServiceTeam> parse(List<String> rawItems) {
        List<ServiceTeam> list = new ArrayList<>();

        for (String b : rawItems) {
            ServiceTeam st = new ServiceTeam();
            String t[] = b.split(" > ");
            String s = t.length>1 ? t[1] : t.length == 1 ? t[0] : "";
            String bu = extractBU(s);
            String ext = extractEXT(s);
            String projectName = extractProjectName(s, bu, ext);

            st.setBu(bu);
            st.setProjectName(projectName);
            st.setExtCode(ext);
            st.setBuDescription(ServiceTeamMaps.BU_DESCRIPTION_MAP.getOrDefault(bu, ""));
            st.setProjectDescription(ServiceTeamMaps.EXT_DESCRIPTION_MAP.getOrDefault(ext, ""));
            if(!projectName.isEmpty()) list.add(st);
        }
        return list;
    }

    private String extractBU(String s) {
        for (String key : ServiceTeamMaps.BU_DESCRIPTION_MAP.keySet()) {
            if (s.startsWith(key)) return key;
        }
        return "";
    }

    private String extractEXT(String s) {
        int idx = s.indexOf("EXT");
        if (idx != -1) return s.substring(idx).trim();
        return "Pending";
    }

    private String extractProjectName(String s, String bu, String ext) {
        int start = bu.isEmpty() ? 0 : s.indexOf(bu) + bu.length() + 1;
        int end = ext.equals("Pending") ? s.length() : s.indexOf(ext) - 1;
        if (start < 0 || end < 0 || start >= end) return "";
        String str = s.substring(start, end).trim();
        str = str.replaceAll("^_+|_+$", "").replace('_', ' ');
        return str;
    }
}