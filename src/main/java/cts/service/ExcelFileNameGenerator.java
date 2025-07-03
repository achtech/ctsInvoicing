package cts.service;

public interface ExcelFileNameGenerator {
    public static final String SHEET_ADJUSTMENT = "Adjustment";
	public static final String SHEET_AJUSTES = "Ajustes";
    public static final String SHEET_SERVICE_HOURS_DETAILS = "Service Hours Details";
	public static final String SHEET_INVOICING_DETAILS = "Invoicing Details";
    public static final String SHEET_HORAS_SERVICIO = "Horas servicio";
	public static final String SHEET_FACTURACION = "Facturación";

	    String generateInputFileName(int year, int month);
	    String generateOutputFileName(int month, int year, String serviceTeam, String directory);
	}
