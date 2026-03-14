// package invoicing.Helper;

// import java.util.Map;

// public class ServiceTeamMaps {
//     public static final Map<String,String> BU_DESCRIPTION_MAP = Map.of(
//         "DT","Digital Technology",
//         "QA","Quality Assurance",
//         "Industry","Industry",
//         "ESS","Enterprise Serv&Solu",
//         "AO","Application Operation",
//         "DE","Digital Experience",
//         "IP","Integration Platform",
//         "TBD","",
//         "BE","Business Engagement",
//         "DSS","DSS – QS&ITO"
//     );

//     public static final Map<String, String> EXT_DESCRIPTION_MAP = Map.ofEntries(
//             Map.entry("EXT-304025-02110", "TH KENA - HPC Marocco"),
//             Map.entry("EXT-303893-02197", "Nautilus 3.0 - Lotto 6 - HPC - FY 2025"),
//             Map.entry("EXT-304218-00001", "C-004499-003 MDM – servizi professionali"),
//             Map.entry("EXT-305784-00006", "DE_Danone_USA"),
//             Map.entry("EXT-304219-00064", "Mulsoft"),
//             Map.entry("EXT-303893-02193", "N3 CAN Lotto 4 - HPC FY 2025"),
//             Map.entry("EXT-304263-02194", "L6 ADM - GR - QA Marocco 2025"),
//             Map.entry("EXT-304292-02010", "Web team FY25 - HPC"),
//             Map.entry("EXT-301864-00040", "AO MSC - HPC MAROCCO"),
//             Map.entry("EXT-304042-00033", "AO VFA - HPC MAROCCO"),
//             Map.entry("EXT-304213-02025", "MA-BMW SCI Direct Costs"),
//             Map.entry("EXT-304042-00034", "Automotive Mobile APP - 2025"),
//             Map.entry("EXT-304069-00027", "Application development 2025 HPC"),
//             Map.entry("EXT-304025-02111", "TH BUS - HPC Marocco"),
//             Map.entry("EXT-304025-02107", "RA 2024-2027 - HPC Marocco")
//     );

// }

package invoicing.Helper;

import java.util.Map;
import java.util.HashMap;

public class ServiceTeamMaps {
    public static final Map<String, String> BU_DESCRIPTION_MAP = Map.of(
        "DT",       "Digital Technology",
        "QA",       "Quality Assurance",
        "Industry", "Industry",
        "ESS",      "Enterprise Serv&Solu",
        "AO",       "Application Operation",
        "DE",       "Digital Experience",
        "IP",       "Integration Platform",
        "TBD",      "",
        "BE",       "Business Engagement",
        "DSS",      "DSS – QS&ITO"
    );

    public static final Map<String, String> CODE_DESCRIPTION_MAP;

    static {
        Map<String, String> map = new HashMap<>();
        // EXT codes
        map.put("EXT-304025-02110", "TH KENA - HPC Marocco");
        map.put("EXT-303893-02197", "Nautilus 3.0 - Lotto 6 - HPC - FY 2025");
        map.put("EXT-304218-00001", "C-004499-003 MDM – servizi professionali");
        map.put("EXT-305784-00006", "DE_Danone_USA");
        map.put("EXT-304219-00064", "Mulsoft");
        map.put("EXT-303893-02193", "N3 CAN Lotto 4 - HPC FY 2025");
        map.put("EXT-304263-02194", "L6 ADM - GR - QA Marocco 2025");
        map.put("EXT-304292-02010", "Web team FY25 - HPC");
        map.put("EXT-301864-00040", "AO MSC - HPC MAROCCO");
        map.put("EXT-304042-00033", "AO VFA - HPC MAROCCO");
        map.put("EXT-304213-02025", "MA-BMW SCI Direct Costs");
        map.put("EXT-304042-00034", "Automotive Mobile APP - 2025");
        map.put("EXT-304069-00027", "Application development 2025 HPC");
        map.put("EXT-304025-02111", "TH BUS - HPC Marocco");
        map.put("EXT-304025-02107", "RA 2024-2027 - HPC Marocco");
        map.put("EXT-303889-00023", "Fideuram ITA");
        map.put("EXT-304080-02141", "DT UNICREDIT ITA");
        map.put("EXT-304291-02005", "DT INWIT");
        map.put("EXT-301864-00066", "GNV");
        // INT codes
        map.put("INT-304434-01064", "DT Sustainability");
        map.put("INT-304434-01066", "Terna Italy");
        // INS codes
        map.put("INS-026696-00003", "CMOR");
        CODE_DESCRIPTION_MAP = map;
    }
}