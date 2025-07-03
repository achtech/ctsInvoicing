package cts.service;

import java.time.LocalDate;
import java.util.Locale;

public interface DateProvider {
    LocalDate getCurrentDate();
    LocalDate getCurrentDate(String directoryPath);
    String getMonthName(LocalDate date, Locale locale);
    String getMonthNameEnglish(LocalDate date);
    String getMonthNameSpanish(LocalDate date);
    int getYear(LocalDate date);
    int getMonthValue(LocalDate date);
}
