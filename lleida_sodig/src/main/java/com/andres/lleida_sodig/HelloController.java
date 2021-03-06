package com.andres.lleida_sodig;

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

import static javafx.application.Application.launch;

public class HelloController {
    @FXML
    public Button btnConexion;
    public Button btnConsultar;
    public Label labelMensaje;
    public DatePicker dateFechaInicial;
    public DatePicker dateFechaFin;
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
    //public  Image img=new Image("E:\\classs\\IntelliJ IDEA\\LLEIDA\\lleida_sodig\\src\\main\\java\\com\\andres\\lleida_sodig\\folder_images_15503.png");
    //file:C:\Users\Andrés Aymacaña\Desktop\sodig.jpg

    public  ImageView imagen=new ImageView();


    public TextField txtUsuario;
    public TextField txtClave;
    public Button btnExcel;
    private HelloModel modelo=new HelloModel();


   /* public static void main(String[] args) {
        imagen.setFitHeight(100);
        imagen.setFitWidth(100);
        imagen.setImage(img);
        launch();

    }*/

    /*@FXML
    public void start() throws Exception {
        imagen.setFitHeight(100);
        imagen.setFitWidth(100);
        imagen.setImage(img);
    }*/


    //ObservableList<Correo> list=FXCollections.observableArrayList();

    //metodo para cargar a lista
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
        table.setMaxSize(500,600);
        table.setMinSize(250,300);
    }

    @FXML
    public void onConexionClickButton(ActionEvent actionEvent) throws IOException {
imagen.setFitHeight(100);
imagen.setFitWidth(100);
//imagen.setImage(img);
      boolean condicion;
        String usuario=txtUsuario.getText();
        String clave=txtClave.getText();

        condicion=modelo.conexion(usuario,clave);
        if (condicion){
            labelMensaje.setText("Ingresó");
            irSecondWindows();
        }else {
            labelMensaje.setText("Intente nuevamente");
        }
    }

    public void onConsultarClickButton(ActionEvent actionEvent) {
        Workbook book = new XSSFWorkbook();
        String fechaInicial=null,fechaFinal=null;
        fechaInicial= dateFechaInicial.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        fechaFinal= dateFechaFin.getValue().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        book =modelo.reportePorFechas(fechaInicial,fechaFinal);
        //labelMensaje.setText(link);

        ObservableList<Correo> list=modelo.enviarLista();
        CargarTabla(list);
    }

    public void irSecondWindows() throws IOException {
        //String mensaje=modelo.obtenerExcel(book);
        //labelMensaje.setText(mensaje);

        FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("Hello2.fxml"));
        Parent root= fxmlLoader.load();
        Hello2 controlador=fxmlLoader.getController();
       //FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("Hello2.fxml"));
        //Scene scene = new Scene(fxmlLoader.load(), 320, 240);
        Scene scene=new Scene(root,400,380);
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
        Stage myStage=(Stage) this.btnConexion.getScene().getWindow();
        myStage.close();
    }

}