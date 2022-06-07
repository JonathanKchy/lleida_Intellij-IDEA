package com.andres.lleida_sodig;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;

public class HelloController {
    private HelloModel modelo=new HelloModel();
    public TextField txtUsuario;
    public TextField txtClave;
    public Button btnExcel;
    @FXML
    private Label welcomeText;

    @FXML
    protected void onHelloButtonClick() {
        welcomeText.setText("Hola mi pez :v");
    }

    public void onExcelClickButton(ActionEvent actionEvent) {
        String usuario=txtUsuario.getText();
        String clave=txtClave.getText();
        welcomeText.setText("Usuario: "+usuario+" Clave: "+clave);
    }
}