package com.example.demo;

import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.time.format.DateTimeFormatter;


public class HelloController {
    @FXML
    public Button btnConexion;
    public Button btnConsultar;
    public Label labelMensaje;

    public Image img=new Image("E:\\classs\\IntelliJ IDEA\\LLEIDA\\lleida_sodig\\src\\main\\java\\com\\andres\\lleida_sodig\\folder_images_15503.png");
    //file:C:\Users\Andrés Aymacaña\Desktop\sodig.jpg

    public ImageView imagen=new ImageView();

    //private HelloModel modelo=new HelloModel();
    public TextField txtUsuario;
    public TextField txtClave;
    public Button btnExcel;


    public void onConexionClickButton(ActionEvent actionEvent) {
    }
}