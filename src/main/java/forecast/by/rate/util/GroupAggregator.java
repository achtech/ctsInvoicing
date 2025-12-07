package forecast.by.rate.util;

// GroupAggregator.java
import java.util.HashMap;
import java.util.Map;

public class GroupAggregator {
    private final Map<String, Double> groupToTotalHoras = new HashMap<>();

    public void addHoras(String group, double horas) {
        groupToTotalHoras.put(group, groupToTotalHoras.getOrDefault(group, 0.0) + horas);
    }

    public Map<String, Double> getAggregates() {
        return new HashMap<>(groupToTotalHoras);
    }
}