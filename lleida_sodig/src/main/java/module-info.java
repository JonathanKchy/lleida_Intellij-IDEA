module com.andres.lleida_sodig {
    requires javafx.controls;
    requires javafx.fxml;

    requires com.dlsc.formsfx;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires java.logging;

    opens com.andres.lleida_sodig to javafx.fxml;
    exports com.andres.lleida_sodig;
}