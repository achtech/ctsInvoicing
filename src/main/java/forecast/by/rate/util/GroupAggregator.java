package forecast.by.rate.util;

// GroupAggregator.java
import java.util.HashMap;
import java.util.Map;

public class GroupAggregator {
    private final Map<String, Map<String, Double>> groupToUserHoras = new HashMap<>();
    private final Map<String, Map<String, Double>> groupToUserFacturacion = new HashMap<>();

    public void add(String group, String user, double horas, double facturacion) {
        // Aggregate Hours
        groupToUserHoras
            .computeIfAbsent(group, k -> new HashMap<>())
            .merge(user, horas, Double::sum);

        // Aggregate Facturacion
        groupToUserFacturacion
            .computeIfAbsent(group, k -> new HashMap<>())
            .merge(user, facturacion, Double::sum);
    }

    public Map<String, Map<String, Double>> getAggregates() {
        // Return hours aggregates (keeping existing signature/behavior for compatibility if needed)
        Map<String, Map<String, Double>> copy = new HashMap<>();
        for (Map.Entry<String, Map<String, Double>> e : groupToUserHoras.entrySet()) {
            copy.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return copy;
    }

    public Map<String, Map<String, Double>> getFacturacionAggregates() {
        Map<String, Map<String, Double>> copy = new HashMap<>();
        for (Map.Entry<String, Map<String, Double>> e : groupToUserFacturacion.entrySet()) {
            copy.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return copy;
    }
}
