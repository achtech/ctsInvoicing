package cts.service.impl;

import cts.service.DateProvider;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.Locale;

public class DefaultDateProvider implements DateProvider {
    //get current date
    public String getCurrentDateTime() {
        LocalDateTime now = LocalDateTime.now();

        // Define the formatter
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm");

        // Format the date
        return now.format(formatter);
    }

    @Override
    public LocalDate getCurrentDate(String path) {
        System.out.println(path);
        String fileName = path.substring(path.lastIndexOf("\\") + 1);
        String[] dates = fileName.split("-");
        int year = Integer.parseInt(dates[0]);
        int month = Integer.parseInt(dates[1]);
        return LocalDate.of(year, month, 1);
    }

    @Override
    public String getMonthName(LocalDate date, Locale locale) {
        return date.getMonth().getDisplayName(TextStyle.FULL, locale);
    }

    @Override
    public String getMonthNameEnglish(LocalDate date) {
        return getMonthName(date, Locale.US);
    }

    @Override
    public String getMonthNameSpanish(LocalDate date) {
        return getMonthName(date, new Locale("es", "ES"));
    }

    @Override
    public int getYear(LocalDate date) {
        return date.getYear();
    }

    @Override
    public int getMonthValue(LocalDate date) {
        return date.getMonthValue();
    }
}

