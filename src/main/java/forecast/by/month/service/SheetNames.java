package forecast.by.month.service;


import java.util.Objects;

public class SheetNames {
    private String horasServicioCurrent;
    private String horasServicioNext;
    private String horasServicioNextNext;
    private String ajustes;
    private String FacturaciónCurrent;
    private String FacturaciónNext;
    private String FacturaciónNextNext;

    private String serviceHoursDetailsCurrent;
    private String serviceHoursDetailsNext;
    private String serviceHoursDetailsNextNext;

    public SheetNames(String horasServicioCurrent) {
        this.horasServicioCurrent = horasServicioCurrent;
    }

    public String getHorasServicioCurrent() {
        return horasServicioCurrent;
    }

    public void setHorasServicioCurrent(String horasServicioCurrent) {
        this.horasServicioCurrent = horasServicioCurrent;
    }

    public String getHorasServicioNext() {
        return horasServicioNext;
    }

    public void setHorasServicioNext(String horasServicioNext) {
        this.horasServicioNext = horasServicioNext;
    }

    public String getHorasServicioNextNext() {
        return horasServicioNextNext;
    }

    public void setHorasServicioNextNext(String horasServicioNextNext) {
        this.horasServicioNextNext = horasServicioNextNext;
    }

    public String getAjustes() {
        return ajustes;
    }

    public void setAjustes(String ajustes) {
        this.ajustes = ajustes;
    }

    public String getFacturaciónCurrent() {
        return FacturaciónCurrent;
    }

    public void setFacturaciónCurrent(String facturaciónCurrent) {
        FacturaciónCurrent = facturaciónCurrent;
    }

    public String getFacturaciónNext() {
        return FacturaciónNext;
    }

    public void setFacturaciónNext(String facturaciónNext) {
        FacturaciónNext = facturaciónNext;
    }

    public String getFacturaciónNextNext() {
        return FacturaciónNextNext;
    }

    public void setFacturaciónNextNext(String facturaciónNextNext) {
        FacturaciónNextNext = facturaciónNextNext;
    }

    public String getServiceHoursDetailsCurrent() {
        return serviceHoursDetailsCurrent;
    }

    public void setServiceHoursDetailsCurrent(String serviceHoursDetailsCurrent) {
        this.serviceHoursDetailsCurrent = serviceHoursDetailsCurrent;
    }

    public String getServiceHoursDetailsNext() {
        return serviceHoursDetailsNext;
    }

    public void setServiceHoursDetailsNext(String serviceHoursDetailsNext) {
        this.serviceHoursDetailsNext = serviceHoursDetailsNext;
    }

    public String getServiceHoursDetailsNextNext() {
        return serviceHoursDetailsNextNext;
    }

    public void setServiceHoursDetailsNextNext(String serviceHoursDetailsNextNext) {
        this.serviceHoursDetailsNextNext = serviceHoursDetailsNextNext;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        SheetNames that = (SheetNames) o;
        return Objects.equals(horasServicioCurrent, that.horasServicioCurrent) && Objects.equals(horasServicioNext, that.horasServicioNext) && Objects.equals(horasServicioNextNext, that.horasServicioNextNext) && Objects.equals(ajustes, that.ajustes) && Objects.equals(FacturaciónCurrent, that.FacturaciónCurrent) && Objects.equals(FacturaciónNext, that.FacturaciónNext) && Objects.equals(FacturaciónNextNext, that.FacturaciónNextNext) && Objects.equals(serviceHoursDetailsCurrent, that.serviceHoursDetailsCurrent) && Objects.equals(serviceHoursDetailsNext, that.serviceHoursDetailsNext) && Objects.equals(serviceHoursDetailsNextNext, that.serviceHoursDetailsNextNext);
    }

    @Override
    public int hashCode() {
        return Objects.hash(horasServicioCurrent, horasServicioNext, horasServicioNextNext, ajustes, FacturaciónCurrent, FacturaciónNext, FacturaciónNextNext, serviceHoursDetailsCurrent, serviceHoursDetailsNext, serviceHoursDetailsNextNext);
    }
}