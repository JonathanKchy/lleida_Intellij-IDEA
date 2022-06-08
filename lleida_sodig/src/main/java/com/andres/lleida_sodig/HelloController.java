package com.andres.lleida_sodig;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;

import java.io.IOException;
import java.util.Date;

public class HelloController {
    private HelloModel modelo=new HelloModel();
    public TextField txtUsuario;
    public TextField txtClave;
    public Button btnExcel;
    @FXML
    private Label welcomeText;

    @FXML
    protected void onHelloButtonClick() throws IOException {
        welcomeText.setText("Hola mi pez :v");
        //modelo.numero();
        modelo.obtenerExcel();
        String fechaInicial=null,fechaFinal=null;
        modelo.reportePorFechas(fechaInicial,fechaFinal);
    }

    public void onExcelClickButton(ActionEvent actionEvent) {
        boolean condicion;
        String usuario=txtUsuario.getText();
        String clave=txtClave.getText();

        condicion=modelo.conexion(usuario,clave);
        if (condicion){
            welcomeText.setText("Ingresó");
        }else {
            welcomeText.setText("Intente nuevamente");
        }
    }
}