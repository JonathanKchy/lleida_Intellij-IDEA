package com.andres.lleida_sodig;

import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.time.format.DateTimeFormatter;

public class Hello2 {
    public Label labellabel;
    public DatePicker dateFechaInicial;
    public DatePicker dateFechaFin;
    public Button btnConsultar;
    public Button btnExcel;
    public TableView<Correo> table;
    public TableColumn<Correo,String> Id;
    public TableColumn<Correo,String> Fecha_Lleida;
    public TableColumn<Correo,String> Fecha_Local;
    public TableColumn<Correo,String> Tipo;
    public TableColumn<Correo,String> Doc_OkKo;
    public TableColumn<Correo,String> Doc_UID;
    public TableColumn<Correo,String> Unidades_Certificadas;
    public TableColumn<Correo,String> Dirección_Origen;
    public TableColumn<Correo,String> Dirección_Destino;
    public TableColumn<Correo,String> Dirección_Cc;
    public TableColumn<Correo,String> Estado;
    public TableColumn<Correo,String> Estado_Aux;
    public TableColumn<Correo,String> Asunto;
    public TableColumn<Correo,String> Doc_Visualizado;
    public TableColumn<Correo,String> Fecha_Visualización;
    public TableColumn<Correo,String> Add_UID;
    public Button btnSofy;

    private HelloModel modelo=new HelloModel();

    static Workbook book = new XSSFWorkbook();

    public void CargarTabla(ObservableList<Correo> list){
        Id.setCellValueFactory(new PropertyValueFactory<Correo,String>("Id"));
        Fecha_Lleida.setCellValueFactory(new PropertyValueFactory<Correo,String>("Fecha_Lleida"));
        Fecha_Local.setCellValueFactory(new PropertyValueFactory<Correo,String>("Fecha_Local"));
        Tipo.setCellValueFactory(new PropertyValueFactory<Correo,String>("Tipo"));
        Doc_OkKo.setCellValueFactory(new PropertyValueFactory<Correo,String>("Doc_OkKo"));
        Doc_UID.setCellValueFactory(new PropertyValueFactory<Correo,String>("Doc_UID"));
        Unidades_Certificadas.setCellValueFactory(new PropertyValueFactory<Correo,String>("Unidades_Certificadas"));
        Dirección_Origen.setCellValueFactory(new PropertyValueFactory<Correo,String>("Dirección_Origen"));
        Dirección_Destino.setCellValueFactory(new PropertyValueFactory<Correo,String>("Dirección_Destino"));
        Dirección_Cc.setCellValueFactory(new PropertyValueFactory<Correo,String>("Dirección_Cc"));
        Estado.setCellValueFactory(new PropertyValueFactory<Correo,String>("Estado"));
        Estado_Aux.setCellValueFactory(new PropertyValueFactory<Correo,String>("Estado_Aux"));
        Asunto.setCellValueFactory(new PropertyValueFactory<Correo,String>("Asunto"));
        Doc_Visualizado.setCellValueFactory(new PropertyValueFactory<Correo,String>("Doc_Visualizado"));
        Fecha_Visualización.setCellValueFactory(new PropertyValueFactory<Correo,String>("Fecha_Visualización"));
        Add_UID.setCellValueFactory(new PropertyValueFactory<Correo,String>("Add_UID"));

        //Correo p=new Correo("2","2","2","2","2","2","2","2","2","2","2","2","2","2","2","2");
        //list.add(p);

        table.setItems(list);
        table.setMaxSize(300,300);
        table.setMinSize(200,200);
        int contadorCorreos=modelo.obtenerContador();
        labellabel.setText("Hay "+contadorCorreos+" correos");
    }

    //cargar visualmete la lista de Sofy
    public void CargarTablaSofy(ObservableList<Correo> list){
        Fecha_Local.setCellValueFactory(new PropertyValueFactory<Correo,String>("Fecha_Local"));
        Doc_OkKo.setCellValueFactory(new PropertyValueFactory<Correo,String>("Doc_OkKo"));
        Dirección_Origen.setCellValueFactory(new PropertyValueFactory<Correo,String>("Dirección_Origen"));
        Dirección_Destino.setCellValueFactory(new PropertyValueFactory<Correo,String>("Dirección_Destino"));
        Doc_Visualizado.setCellValueFactory(new PropertyValueFactory<Correo,String>("Doc_Visualizado"));
        //falta columna del SI
        //Correo p=new Correo("2","2","2","2","2","2","2","2","2","2","2","2","2","2","2","2");
        //list.add(p);

        table.setItems(list);
        table.setMaxSize(300,300);
        table.setMinSize(200,200);
        int contadorCorreos=modelo.obtenerContador();
        labellabel.setText("Hay "+contadorCorreos+" correos");
    }

    public void closeWindows() throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("hello-view.fxml"));
        Parent root= fxmlLoader.load();

        Scene scene=new Scene(root);
        Stage stage = new Stage();
        stage.setTitle("Hello!");
        stage.setScene(scene);
        stage.show();
        //stage.setOnCloseRequest(e -> controlador.closeWindows());
        Stage myStage=(Stage) this.labellabel.getScene().getWindow();
        myStage.close();
    }

    public void onConsultarClickButton(ActionEvent actionEvent) {
        btnExcel.setDisable(false);
        table.getItems().clear();
        String fechaInicial=null,fechaFinal=null;
        fechaInicial= dateFechaInicial.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        fechaFinal= dateFechaFin.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        book=null;
        book =modelo.reportePorFechas(fechaInicial,fechaFinal);
        //labelMensaje.setText(link);

        ObservableList<Correo> list=modelo.enviarLista();
        CargarTabla(list);
    }

    public void onGenerarClickButton(ActionEvent actionEvent) throws IOException {
        btnExcel.setDisable(true);
        String mensaje=modelo.obtenerExcel(book);
        labellabel.setText(mensaje);

       /* FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("Hello2.fxml"));
        Parent root= fxmlLoader.load();
        Hello2 controlador=fxmlLoader.getController();
        //FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("Hello2.fxml"));
        //Scene scene = new Scene(fxmlLoader.load(), 320, 240);
        Scene scene=new Scene(root);
        Stage stage = new Stage();
        stage.setTitle("Hello2!");
        stage.setScene(scene);
        stage.show();
        stage.setOnCloseRequest(e -> {
            try {
                controlador.closeWindows();
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
        });
        Stage myStage=(Stage) this.labellabel.getScene().getWindow();
        myStage.close();*/
    }

    public void onConsultarClickButtonSofy(ActionEvent actionEvent) {
        btnExcel.setDisable(false);
        table.getItems().clear();
        String fechaInicial=null,fechaFinal=null;
        fechaInicial= dateFechaInicial.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        fechaFinal= dateFechaFin.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        book=null;
        book =modelo.reportePorFechasSofy(fechaInicial,fechaFinal);
        //labelMensaje.setText(link);

        ObservableList<Correo> list=modelo.enviarLista();
        CargarTablaSofy(list);
    }
}
