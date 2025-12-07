package forecast.by.month.service;

import java.util.Map;

public class RateTable {
    private static final java.util.Map<Double, String> rateMap;
    private static final java.util.Map<Double, Double> exactRateMap;

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

        exactRateMap = new java.util.HashMap<>();
        exactRateMap.put(6.40, 6.39768);
        exactRateMap.put(9.36, 9.36472);
        exactRateMap.put(9.55, 9.55016);
//        exactRateMap.put(9.92, "APP-NSKL-6");
        exactRateMap.put(10.2, 10.1992);
        exactRateMap.put(12.52, 12.5172);
        exactRateMap.put(13.54, 13.53712);
        exactRateMap.put(14.09, 14.09344);
        exactRateMap.put(14.37, 14.3716);
        exactRateMap.put(16.60, 16.59688);
        exactRateMap.put(17.15, 17.1532);
        exactRateMap.put(18.08, 18.0804);
  //      exactRateMap.put(19.93, "APP-TEST-9");
        exactRateMap.put(20.40, 20.3984);
/*
        exactRateMap.put(20.77, "APP-DEV-9");
        exactRateMap.put(21.33, "APP-NSKL-9");
        exactRateMap.put(23.83, "APP-TEST-10");
*/
        exactRateMap.put(24.29, 24.29264);
        exactRateMap.put(24.57, 24.5708);
        exactRateMap.put(25.03, 25.0344);
//        exactRateMap.put(25.96, "APP-PM-10");
//        exactRateMap.put(27.45, "APP-DEV-11");
//        exactRateMap.put(29.67, "APP-DA-11");
        exactRateMap.put(30.60, 30.5976);
//        exactRateMap.put(38.02, "APP-DEV-12");
//        exactRateMap.put(38.48, "APP-PM-12");
//        exactRateMap.put(38.94, "APP-DA-12");
//        exactRateMap.put(41.72, "APP-PM-13");
//        exactRateMap.put(42.65, "APP-DEV-13");
//        exactRateMap.put(46.36, "APP-DA-13");

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

    public static double exactRate(double approximateRate) {
        double tolerance = 0.01; // Adjust tolerance as needed
        for (Map.Entry<Double, Double> entry : exactRateMap.entrySet()) {
            if (Math.abs(entry.getKey() - approximateRate) < tolerance) {
                return entry.getValue();
            }
        }
        return approximateRate;
    }

}