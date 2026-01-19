package forecast.by.rate.util;

// GroupAggregator.java
import java.util.HashMap;
import java.util.Map;

public class GroupAggregator {
    private final Map<String, Map<String, Double>> groupToUserHoras = new HashMap<>();

    public void addHoras(String group, String user, double horas) {
        Map<String, Double> userMap = groupToUserHoras.get(group);
        if (userMap == null) {
            userMap = new HashMap<>();
            groupToUserHoras.put(group, userMap);
        }
        Double current = userMap.get(user);
        if (current == null) {
            current = 0.0;
        }
        userMap.put(user, current + horas);
    }

    public Map<String, Map<String, Double>> getAggregates() {
        Map<String, Map<String, Double>> copy = new HashMap<>();
        for (Map.Entry<String, Map<String, Double>> e : groupToUserHoras.entrySet()) {
            copy.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return copy;
    }
}
