package forecast.by.month.service;

public interface ExcelFileNameGenerator {
    String SHEET_AJUSTES = "Ajustes";
    String SHEET_SERVICE_HOURS_DETAILS = "Service Hours Details";
    String SHEET_HORAS_SERVICIO = "Horas servicio";
    String SHEET_FACTURACIÓN = "Facturación";

    String generateOutputFileName(int month, int year, String serviceTeam, String directory);
}
