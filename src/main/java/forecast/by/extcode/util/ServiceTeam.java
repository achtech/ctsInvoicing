package forecast.by.extcode.util;

import org.apache.poi.ss.usermodel.CellStyle;

public class ServiceTeam {
    private String bu;
    private String projectName;
    private String extCode;
    private String cost;
    private CellStyle style;

    private String projectDescription;
    private String buDescription;

    public CellStyle getStyle() {
        return style;
    }

    public void setStyle(CellStyle style) {
        this.style = style;
    }

    public String getBu() { return bu; }
    public void setBu(String bu) { this.bu = bu; }

    public String getProjectName() { return projectName; }
    public void setProjectName(String projectName) { this.projectName = projectName; }

    public String getExtCode() { return extCode; }
    public void setExtCode(String extCode) { this.extCode = extCode; }

    public String getCost() { return cost; }
    public void setCost(String cost) { this.cost = cost; }

    public String getProjectDescription() { return projectDescription; }
    public void setProjectDescription(String projectDescription) { this.projectDescription = projectDescription; }

    public String getBuDescription() { return buDescription; }
    public void setBuDescription(String buDescription) { this.buDescription = buDescription; }
}
