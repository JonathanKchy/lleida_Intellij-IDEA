package com.andres.lleida_sodig;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.stage.Stage;

import java.io.IOException;

public class Hello2 {
    public Label labellabel;

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
}
