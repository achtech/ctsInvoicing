package forecast.by.util;

import java.util.Map;

public class ServiceTeamMaps {
    public static final Map<String,String> BU_DESCRIPTION_MAP = Map.of(
        "DT","Digital Technology",
        "QA","Quality Assurance",
        "Industry","Industry",
        "ESS","Enterprise Serv&Solu",
        "AO","Application Operation",
        "DE","Digital Experience",
        "IP","Integration Platform",
        "TBD","",
        "BE","Business Engagement",
        "DSS","DSS – QS&ITO"
    );

    public static final Map<String, String> EXT_DESCRIPTION_MAP = Map.ofEntries(
            Map.entry("EXT-304025-02110", "TH KENA - HPC Marocco"),
            Map.entry("EXT-303893-02197", "Nautilus 3.0 - Lotto 6 - HPC - FY 2025"),
            Map.entry("EXT-304218-00001", "C-004499-003 MDM – servizi professionali"),
            Map.entry("EXT-305784-00006", "DE_Danone_USA"),
            Map.entry("EXT-304219-00064", "Mulsoft"),
            Map.entry("EXT-303893-02193", "N3 CAN Lotto 4 - HPC FY 2025"),
            Map.entry("EXT-304263-02194", "L6 ADM - GR - QA Marocco 2025"),
            Map.entry("EXT-304292-02010", "Web team FY25 - HPC"),
            Map.entry("EXT-301864-00040", "AO MSC - HPC MAROCCO"),
            Map.entry("EXT-304042-00033", "AO VFA - HPC MAROCCO"),
            Map.entry("EXT-304213-02025", "MA-BMW SCI Direct Costs"),
            Map.entry("EXT-304042-00034", "Automotive Mobile APP - 2025"),
            Map.entry("EXT-304069-00027", "Application development 2025 HPC"),
            Map.entry("EXT-304025-02111", "TH BUS - HPC Marocco"),
            Map.entry("EXT-304025-02107", "RA 2024-2027 - HPC Marocco")
    );

}
