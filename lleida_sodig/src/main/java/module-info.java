module com.andres.lleida_sodig {
    requires javafx.controls;
    requires javafx.fxml;

    requires com.dlsc.formsfx;

    opens com.andres.lleida_sodig to javafx.fxml;
    exports com.andres.lleida_sodig;
}