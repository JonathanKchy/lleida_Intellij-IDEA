package com.andres.lleida_sodig;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRotY;

import javax.net.ssl.HttpsURLConnection;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.ProtocolException;
import java.net.URL;
import java.util.Date;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;

public class HelloModel {

    static String mail_id, mail_date,fecha_andre, mail_type, file_doc_model, file_uid, unidades_certificadas, mail_from, mail_to,direccion_CC="correo@certificado.lleida.net", gstatus, gstatus_aux, mail_subj, add_id, add_displaydate, add_uid;
    static String user,password,link="";
    static Workbook book = new XSSFWorkbook();

    static ObservableList<Correo> list= FXCollections.observableArrayList();
//metodo enviar lista

    public ObservableList<Correo> enviarLista(){
        return list;
    }


    //metodo para validar el usuario y la clave
    public boolean conexion(String usuario,String clave){
    boolean bool=false;
    user=usuario;
    password=clave;
        try

        {
            URL url = new URL("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=get_default_settings&user="+usuario+"&password="+clave);
            HttpsURLConnection conn = (HttpsURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.connect();
            int responseCod = conn.getResponseCode();
            if (responseCod != 200) {
                bool=false;
            } else {
                Scanner scanner=new Scanner(url.openStream());
                while (scanner.hasNext()) {
                    String linea = scanner.nextLine();
                    if (linea.contains("<status>")) {
                        int tamano = linea.length();
                        int fin = tamano - 9;
                        int status = Integer.parseInt(linea.substring(8, fin));
                        System.out.println(status);
                        if (status==100){
                            bool=true;
                        }else {
                            bool=false;
                        }
                    }
                }

            }
        }catch(Exception e){

        }
    return  bool;
    }


    //metodo para obtener reporte por rango de fechas
    public Workbook reportePorFechas(String fechaInicial,String fechaFinal){
    String fecha=null;
       link="https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user="+user+"&password="+password+"&mail_date_min="+fechaInicial+"070000&mail_date_max="+fechaFinal+"070000";
       int contador = 0, numeroCelda = 0;

        // Creamos el libro de trabajo de Excel formato OOXML
        book = new XSSFWorkbook();
        // La hoja donde pondremos los datos
        Sheet sheet = book.createSheet("Hola Java");
        //creamos una fila
        Row row = sheet.createRow(contador);
        row.createCell(0).setCellValue("Id");
        row.createCell(1).setCellValue("Fecha Lleida");
        row.createCell(2).setCellValue("Fecha Local");
        row.createCell(3).setCellValue("Tipo");
        row.createCell(4).setCellValue("Doc_OkKo");
        row.createCell(5).setCellValue("Doc_UID");
        row.createCell(6).setCellValue("Unidades Certificadas");
        row.createCell(7).setCellValue("Dirección Origen");
        row.createCell(8).setCellValue("Dirección Destino");
        row.createCell(9).setCellValue("Dirección Cc");
        row.createCell(10).setCellValue("Estado");
        row.createCell(11).setCellValue("Estado Aux");
        row.createCell(12).setCellValue("Asunto");
        row.createCell(13).setCellValue("Doc_Visualizado");
        row.createCell(14).setCellValue("Fecha y hora de visualización");
        row.createCell(15).setCellValue("Add_UID");

        int dia_anterior = 1;

        try {

            URL url= new URL(link);
            HttpsURLConnection conn=(HttpsURLConnection)url.openConnection();
            conn.setRequestMethod("GET");
            conn.connect();

            int responseCod= conn.getResponseCode();
            if (responseCod!=200) {
                throw new RuntimeException("ocurrio un error: "+responseCod);
            }else  {
                StringBuilder informationString=new StringBuilder();
                Scanner scanner=new Scanner(url.openStream());
                while (scanner.hasNext()) {
                    String linea=scanner.nextLine();

                    if (linea.contains("<mail_id>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        mail_id=linea.substring(9, fin);
                        informationString.append("mail_id: "+mail_id);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<mail_date>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        int andre_dia, andre_hora,mes_andre;
                        String andre_dia_string,andre_hora_string="01";
                        mes_andre = Integer.parseInt(linea.substring(15,17));
                        andre_dia = Integer.parseInt(linea.substring(17,19));
                        andre_hora = Integer.parseInt(linea.substring(19,21));

                        if (dia_anterior <=andre_dia)
                        {
                            dia_anterior = andre_dia;
                        }
                        switch (andre_hora)
                        {
                            case 23:
                                andre_hora_string = "16";
                                break;
                            case 22:
                                andre_hora_string = "15";
                                break;
                            case 21:
                                andre_hora_string = "14";
                                break;
                            case 20:
                                andre_hora_string = "13";
                                break;
                            case 19:
                                andre_hora_string = "12";
                                break;
                            case 18:
                                andre_hora_string = "11";
                                break;
                            case 17:
                                andre_hora_string = "10";
                                break;
                            case 16:
                                andre_hora_string = "09";
                                break;
                            case 15:
                                andre_hora_string = "08";
                                break;
                            case 14:
                                andre_hora_string = "07";
                                break;
                            case 13:
                                andre_hora_string = "06";
                                break;
                            case 12:
                                andre_hora_string = "05";
                                break;
                            case 11:
                                andre_hora_string = "04";
                                break;
                            case 10:
                                andre_hora_string = "03";
                                break;
                            case 9:
                                andre_hora_string = "02";
                                break;
                            case 8:
                                andre_hora_string = "01";
                                break ;
                            case 7:
                                andre_hora_string = "00";
                                break;
                            case 6:
                                andre_hora_string = "23";
                                break;
                            case 5:
                                andre_hora_string = "22";
                                break;
                            case 4:
                                andre_hora_string = "21";
                                break;
                            case 3:
                                andre_hora_string = "20";
                                break;
                            case 2:
                                andre_hora_string = "19";
                                break;
                            case 1:
                                andre_hora_string = "18";
                                break;
                            case 0:
                                andre_hora_string = "17";
                                break;
                            default:
                                break;
                        }
                        if (andre_hora<7)
                        {
                            andre_dia = andre_dia-1;
                            if (andre_dia==0)
                            {
                                andre_dia = dia_anterior;
                                mes_andre = mes_andre-1;
                            }

                        }
                        fecha_andre = andre_dia + "/" +mes_andre + "/" + linea.substring(11,15)+ " " + andre_hora_string + ":" + linea.substring(21,23)+ ":" + linea.substring(23,25);
                        mail_date=linea.substring(17,19)+"/"+linea.substring(15,17)+"/"+linea.substring(11,15)+" "+linea.substring(19,21)+":"+linea.substring(21,23)+":"+linea.substring(23,25);

                        informationString.append("mail_date: "+mail_date);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<mail_type>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        mail_type=linea.substring(11, fin);

                        informationString.append("mail_date: "+mail_type);
                        informationString.append("\n");
                    } else if(linea.contains("<credits>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        unidades_certificadas=linea.substring(9, fin);

                        if (Integer.parseInt(unidades_certificadas)==14){
                            unidades_certificadas="2.00";
                        }else {
                            unidades_certificadas="1.00";
                        }
                        informationString.append("unidades_certificadas: "+unidades_certificadas);

                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<file_doc_model>")) {
                        int tamano=linea.length();
                        int fin=tamano-17;
                        file_doc_model=linea.substring(16, fin);
                        informationString.append("file_doc_model: "+file_doc_model);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<file_uid>")) {
                        int tamano=linea.length();
                        int fin=tamano-11;
                        file_uid=linea.substring(10, fin);
                        informationString.append("file_uid: "+file_uid);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<mail_from>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        mail_from=linea.substring(11, fin);
                        informationString.append("mail_from: "+mail_from);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<mail_to>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        mail_to=linea.substring(9, fin);
                        informationString.append("mail_to: "+mail_to);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<gstatus>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        gstatus=linea.substring(9, fin);
                        informationString.append("gstatus: "+gstatus);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<gstatus_aux>")) {
                        int tamano=linea.length();
                        int fin=tamano-14;
                        gstatus_aux=linea.substring(13, fin);
                        informationString.append("gstatus_aux: "+gstatus_aux);
                        informationString.append("\n");
                    }else if(linea.contains("<gstatus_aux/>")) {
                        gstatus_aux="";
                        informationString.append("gstatus_aux: "+gstatus_aux);                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<mail_subj>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        mail_subj=linea.substring(11, fin);
                        informationString.append("mail_subj: "+mail_subj);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<add_id>")) {
                        int tamano=linea.length();
                        int fin=tamano-9;
                        add_id=linea.substring(8, fin);
                        add_uid = "E" + add_id + "-R";
                        add_id = "Displayed";
                        informationString.append("add_id: "+add_id);
                        informationString.append("\n");
                    }
                    //+++++++++++++++++++++++++++++
                    else if(linea.contains("<add_displaydate>")) {
                        int tamano=linea.length();
                        int fin=tamano-18;
                        add_displaydate=linea.substring(23,25)+"/"+linea.substring(21,23)+"/"+linea.substring(17,21)+" "+linea.substring(25,27)+":"+linea.substring(27,29)+":"+linea.substring(29,31);

                        informationString.append("add_displaydate: "+add_displaydate);
                        informationString.append("\n");
                    }
                    else if(linea.contains("</pdf_row>")) {

                        informationString.append("\n");
                        contador++;
                        row=sheet.createRow(contador);

                        row.createCell(numeroCelda).setCellValue(mail_id);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_date);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(fecha_andre);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_type);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(file_doc_model);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(file_uid);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(unidades_certificadas);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_from);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_to);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(direccion_CC);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(gstatus);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(gstatus_aux);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_subj);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(add_id);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(add_displaydate);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(add_uid);
                        numeroCelda++;

                        Correo p=new Correo(mail_id,mail_date,fecha_andre,mail_type,file_doc_model,file_uid,unidades_certificadas,mail_from,mail_to,direccion_CC,gstatus,gstatus_aux,mail_subj,add_id,add_displaydate,add_uid);
                        list.add(p);
                        //seteo de variables
                        mail_date=mail_date=unidades_certificadas=mail_type=fecha_andre=file_doc_model=file_uid=mail_from=mail_to=gstatus=gstatus_aux=mail_subj=add_id=add_uid=add_displaydate="";
                        //nueva fila
                        numeroCelda=0;
                    }

                }
                scanner.close();
                //System.out.println(informationString);
                System.out.println(contador);
            }
        } catch (Exception e) {
            System.out.println("error: ");
        }
        //CrearExcel(book);
        //System.out.println("hecho: ");
        return book;
    }

    //metodo para obtener Excel
    public String obtenerExcel(Workbook book) {
        String mensaje="no se puedo descargar";
        // Creamos el estilo paga las celdas del encabezado
        CellStyle style = book.createCellStyle();
        // Indicamos que tendra un fondo azul aqua
        // con patron solido del color indicado
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        try {
            // Creamos el flujo de salida de datos,
            // apuntando al archivo donde queremos
            // almacenar el libro de Excel
            String desktopPath = System.getProperty("user.home") + "/Desktop";
            System.out.println(desktopPath.replace("\\", "/"));
            FileOutputStream fileout=new FileOutputStream(desktopPath.replace("\\", "/")+"/Excel.xlsx");
            try {
                // Almacenamos el libro de
                // Excel via ese
                // flujo de datos
                book.write(fileout);
                mensaje="El excel está en el escritorio";
            } catch (IOException ex) {
                Logger.getLogger(HelloModel.class.getName()).log(Level.SEVERE, null, ex);
            }
            try {
                // Cerramos el libro para concluir operaciones
                fileout.close();
                //  LOGGER.log(Level.INFO, "Archivo creado existosamente en {0}", fil.getAbsolutePath());
            } catch (IOException ex) {
                Logger.getLogger(HelloModel.class.getName()).log(Level.SEVERE, null, ex);
            }

        } catch (FileNotFoundException ex) {
            Logger.getLogger(HelloModel.class.getName()).log(Level.SEVERE, null, ex);
        }
        return mensaje;
    }
}