package com.andres.lleida_sodig;

import javafx.scene.control.TableColumn;

public class Correo {

    private String Id;
    private String Fecha_Lleida;
    private String Fecha_Local;

    private String Tipo;
    private String Doc_OkKo;
    private String Doc_UID;
    private String Unidades_Certificadas;
    private String Dirección_Origen;
    private String Dirección_Destino;
    private String Dirección_Cc;
    private String Estado;
    private String Estado_Aux;
    private String Asunto;
    private String Doc_Visualizado;
    private String Fecha_Visualización;
    private String Add_UID;

    public Correo(String id, String fecha_Lleida, String fecha_Local, String tipo, String doc_OkKo, String doc_UID, String unidades_Certificadas, String dirección_Origen, String dirección_Destino, String dirección_Cc, String estado, String estado_Aux, String asunto, String doc_Visualizado, String fecha_Visualización, String add_UID) {
        Id = id;
        Fecha_Lleida = fecha_Lleida;
        Fecha_Local = fecha_Local;
        Tipo = tipo;
        Doc_OkKo = doc_OkKo;
        Doc_UID = doc_UID;
        Unidades_Certificadas = unidades_Certificadas;
        Dirección_Origen = dirección_Origen;
        Dirección_Destino = dirección_Destino;
        Dirección_Cc = dirección_Cc;
        Estado = estado;
        Estado_Aux = estado_Aux;
        Asunto = asunto;
        Doc_Visualizado = doc_Visualizado;
        Fecha_Visualización = fecha_Visualización;
        Add_UID = add_UID;
    }

    public String getId() {
        return Id;
    }

    public String getFecha_Lleida() {
        return Fecha_Lleida;
    }

    public String getFecha_Local() {
        return Fecha_Local;
    }

    public String getTipo() {
        return Tipo;
    }

    public String getDoc_OkKo() {
        return Doc_OkKo;
    }

    public String getDoc_UID() {
        return Doc_UID;
    }

    public String getUnidades_Certificadas() {
        return Unidades_Certificadas;
    }

    public String getDirección_Origen() {
        return Dirección_Origen;
    }

    public String getDirección_Destino() {
        return Dirección_Destino;
    }

    public String getDirección_Cc() {
        return Dirección_Cc;
    }

    public String getEstado() {
        return Estado;
    }

    public String getEstado_Aux() {
        return Estado_Aux;
    }

    public String getAsunto() {
        return Asunto;
    }

    public String getDoc_Visualizado() {
        return Doc_Visualizado;
    }

    public String getFecha_Visualización() {
        return Fecha_Visualización;
    }

    public String getAdd_UID() {
        return Add_UID;
    }
}
