package forecast.by.month.service.Noura.Model;

public class Ticket {
    private String ticketName;
    private String prDate;
    private String prResponsible;
    private String prDeveloper;

    public Ticket() {}

    public Ticket(String ticketName, String prDate, String prResponsible, String prDeveloper) {
        this.ticketName = ticketName;
        this.prDate = prDate;
        this.prResponsible = prResponsible;
        this.prDeveloper = prDeveloper;
    }

    public String getTicketName() {
        return ticketName;
    }

    public void setTicketName(String ticketName) {
        this.ticketName = ticketName;
    }

    public String getPrDate() {
        return prDate;
    }

    public void setPrDate(String prDate) {
        this.prDate = prDate;
    }

    public String getPrResponsible() {
        return prResponsible;
    }

    public void setPrResponsible(String prResponsible) {
        this.prResponsible = prResponsible;
    }

    public String getPrDeveloper() {
        return prDeveloper;
    }

    public void setPrDeveloper(String prDeveloper) {
        this.prDeveloper = prDeveloper;
    }

    @Override
    public String toString() {
        return "Ticket{" +
                "ticketName='" + ticketName + '\'' +
                ", prDate='" + prDate + '\'' +
                ", prResponsible='" + prResponsible + '\'' +
                ", prDeveloper='" + prDeveloper + '\'' +
                '}';
    }
}
