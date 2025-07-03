package cts.service;

import java.util.Map;

public class RateTable {
    private static java.util.Map<Double, String> rateMap;

    static {
        rateMap = new java.util.HashMap<>();
        rateMap.put(6.40, "APP-DEV-5");
        rateMap.put(9.36, "APP-DEV-6");
        rateMap.put(9.55, "APP-TEST-6");
        rateMap.put(9.92, "APP-NSKL-6");
        rateMap.put(10.2, "APP-DA-6");
        rateMap.put(12.52, "APP-TEST-7");
        rateMap.put(13.54, "APP-DEV-7");
        rateMap.put(14.09, "APP-DA-7");
        rateMap.put(14.37, "APP-NSKL-7");
        rateMap.put(16.60, "APP-TEST-8");
        rateMap.put(17.15, "APP-DEV-8");
        rateMap.put(18.08, "APP-DA-8 / APP-NSKL-8");
        rateMap.put(19.93, "APP-TEST-9");
        rateMap.put(20.40, "APP-DA-9");
        rateMap.put(20.77, "APP-DEV-9");
        rateMap.put(21.33, "APP-NSKL-9");
        rateMap.put(23.83, "APP-TEST-10");
        rateMap.put(24.29, "APP-DEV-10");
        rateMap.put(24.57, "APP-DA-10");
        rateMap.put(25.03, "APP-NSKL-10");
        rateMap.put(25.96, "APP-PM-10");
        rateMap.put(27.45, "APP-DEV-11");
        rateMap.put(29.67, "APP-DA-11");
        rateMap.put(30.60, "APP-PM-11");
        rateMap.put(38.02, "APP-DEV-12");
        rateMap.put(38.48, "APP-PM-12");
        rateMap.put(38.94, "APP-DA-12");
        rateMap.put(41.72, "APP-PM-13");
        rateMap.put(42.65, "APP-DEV-13");
        rateMap.put(46.36, "APP-DA-13");
    }

    public static String getCategory(double approximateRate) {
        double tolerance = 0.01; // Adjust tolerance as needed
        for (Map.Entry<Double, String> entry : rateMap.entrySet()) {
            if (Math.abs(entry.getKey() - approximateRate) < tolerance) {
                return entry.getValue();
            }
        }
        return "Category not found";
    }
}