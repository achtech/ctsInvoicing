package forecast.by.util;
import java.math.BigDecimal;

public class CogsRecord {

    private String groupId;
    private BigDecimal fy25;
    private BigDecimal fy26;

    public CogsRecord(String groupId, BigDecimal fy25, BigDecimal fy26) {
        this.groupId = groupId;
        this.fy25 = fy25;
        this.fy26 = fy26;
    }

    public String getGroupId() { return groupId; }
    public BigDecimal getFy25() { return fy25; }
    public BigDecimal getFy26() { return fy26; }
}
