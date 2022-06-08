package com.andres.lleida_sodig;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class HelloController {
    @FXML
    public Button btnConexion;
    public Button btnConsultar;
    public Label labelMensaje;
    public DatePicker dateFechaInicial;
    public DatePicker dateFechaFin;
    private HelloModel modelo=new HelloModel();
    public TextField txtUsuario;
    public TextField txtClave;
    public Button btnExcel;

    static Workbook book = new XSSFWorkbook();

    @FXML
    public void onConexionClickButton(ActionEvent actionEvent) {
        boolean condicion;
        String usuario=txtUsuario.getText();
        String clave=txtClave.getText();

        condicion=modelo.conexion(usuario,clave);
        if (condicion){
            labelMensaje.setText("Ingresó");
        }else {
            labelMensaje.setText("Intente nuevamente");
        }
    }

    public void onConsultarClickButton(ActionEvent actionEvent) {
        String fechaInicial=null,fechaFinal=null;
        fechaInicial= dateFechaInicial.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        fechaFinal= dateFechaFin.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        book =modelo.reportePorFechas(fechaInicial,fechaFinal);
        //labelMensaje.setText(link);
    }

    public void onGenerarClickButton(ActionEvent actionEvent) {
        String mensaje=modelo.obtenerExcel(book);
        labelMensaje.setText(mensaje);
    }
}