package forecast.by.util;

// GroupAggregator.java
import java.util.HashMap;
import java.util.Map;

public class GroupAggregator {
    private final Map<String, Map<String, Double>> groupToUserHoras = new HashMap<>();
    private final Map<String, Map<String, Double>> groupToUserCost = new HashMap<>();

    public void add(String group, String user, double horas, double cost) {
        // Aggregate Hours
        groupToUserHoras
            .computeIfAbsent(group, k -> new HashMap<>())
            .merge(user, horas, Double::sum);

        // Aggregate Cost
        groupToUserCost
            .computeIfAbsent(group, k -> new HashMap<>())
            .merge(user, cost, Double::sum);
    }

    public Map<String, Map<String, Double>> getAggregates() {
        // Return hours aggregates
        Map<String, Map<String, Double>> copy = new HashMap<>();
        for (Map.Entry<String, Map<String, Double>> e : groupToUserHoras.entrySet()) {
            copy.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return copy;
    }

    public Map<String, Map<String, Double>> getCostAggregates() {
        Map<String, Map<String, Double>> copy = new HashMap<>();
        for (Map.Entry<String, Map<String, Double>> e : groupToUserCost.entrySet()) {
            copy.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return copy;
    }
}
