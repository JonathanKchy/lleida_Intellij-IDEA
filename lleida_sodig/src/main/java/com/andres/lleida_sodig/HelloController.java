package com.andres.lleida_sodig;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
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
    public TableView<Product> table;
    public TableColumn<Product,String> tcID;
    public TableColumn<Product,String> tcDate;
    public TableColumn<Product,String> tcDate2;
    private HelloModel modelo=new HelloModel();
    public TextField txtUsuario;
    public TextField txtClave;
    public Button btnExcel;

    static Workbook book = new XSSFWorkbook();

    ObservableList<Product> list=FXCollections.observableArrayList();


    @FXML
    public void onConexionClickButton(ActionEvent actionEvent) {
        //
        tcID.setCellValueFactory(new PropertyValueFactory<Product,String>("name"));
        tcDate.setCellValueFactory(new PropertyValueFactory<Product,String>("age"));
        tcDate2.setCellValueFactory(new PropertyValueFactory<Product,String>("animal"));
        Product p=new Product("2","2","2");
        list.add(p);
        table.setItems(list);

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