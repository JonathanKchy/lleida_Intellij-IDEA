package com.andres.lleida_sodig;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRotY;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlurEffect;

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

    static String mail_id, mail_date,fecha_Ecuador, mail_type, file_doc_model, file_uid, unidades_certificadas, mail_from, mail_to,direccion_CC="correo@certificado.lleida.net", gstatus, gstatus_aux, mail_subj, add_id, add_displaydate, add_uid;
    static String user,password,link="";
    static Workbook book = new XSSFWorkbook();
    public int contador=0;
    public ObservableList<Correo> list= FXCollections.observableArrayList();
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
        list.clear();
    String fecha=null;
       link="https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user="+user+"&password="+password+"&mail_date_min="+fechaInicial+"070000&mail_date_max="+fechaFinal+"070000";
       contador = 0;
       int numeroCelda = 0;

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
        row.createCell(7).setCellValue("Direcci??n Origen");
        row.createCell(8).setCellValue("Direcci??n Destino");
        row.createCell(9).setCellValue("Direcci??n Cc");
        row.createCell(10).setCellValue("Estado");
        row.createCell(11).setCellValue("Estado Aux");
        row.createCell(12).setCellValue("Asunto");
        row.createCell(13).setCellValue("Doc_Visualizado");
        row.createCell(14).setCellValue("Fecha y hora de visualizaci??n");
        row.createCell(15).setCellValue("Add_UID");

        //esta variable servir?? para almacenar el n??mero del d??a anterior, sirve para poder cambiar
        //la fecha que nos entrega lleida a la fecha de Ecuador ya que el servidor de lleida est??
        //adelantado en relaci??n de  nosotros por 7 horas
        int dia_anterior = 1;
        String[] newStr = null;
        String correoAnterior = "";

        try {

            URL url= new URL(link);
            HttpsURLConnection conn=(HttpsURLConnection)url.openConnection();
            conn.setRequestMethod("GET");
            conn.connect();
            System.out.println(link);
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
                        //int tamano=linea.length();
                        //int fin=tamano-12;

                        //dia_Ecuador, permite capturar el valor del d??a del correo de lleida
                        //hora_Ecuador, permite capturar el valor de la hora del correo de lleida
                        //mes_Ecuador, permite capturar el valor del mes del correo de lleida
                        int dia_Ecuador, hora_Ecuador,mes_Ecuador;
                        // hora_Ecuador_string, permite capturar la hora del correo en formato string,
                        //se setea en 01 solo por evitar el try catch
                        String hora_Ecuador_string="01";
                        mes_Ecuador = Integer.parseInt(linea.substring(15,17));
                        dia_Ecuador = Integer.parseInt(linea.substring(17,19));
                        hora_Ecuador = Integer.parseInt(linea.substring(19,21));
                        //se compara si el d??a anterior que se seteo en 1, es menosr que el d??a capturado
                        //desde lleida, as?? podemos setear el d??a de Ecuador. Ejemplo si el correo desde
                        //lleida me llega con fecha 01/03/2022 05h:30m:20s yo debo tener el d??a anterior
                        //se el 30 o 31 y poner fecha de Ecuador como 30/02/2022 22h:30m:20s
                        if (dia_anterior <=dia_Ecuador)
                        {
                            dia_anterior = dia_Ecuador;
                        }
                        //como se que la hora de lleida tiene 7 horas en adelanto, se debe setear la hora de Ecuador
                        switch (hora_Ecuador)
                        {
                            case 23:
                                hora_Ecuador_string = "16";
                                break;
                            case 22:
                                hora_Ecuador_string = "15";
                                break;
                            case 21:
                                hora_Ecuador_string = "14";
                                break;
                            case 20:
                                hora_Ecuador_string = "13";
                                break;
                            case 19:
                                hora_Ecuador_string = "12";
                                break;
                            case 18:
                                hora_Ecuador_string = "11";
                                break;
                            case 17:
                                hora_Ecuador_string = "10";
                                break;
                            case 16:
                                hora_Ecuador_string = "09";
                                break;
                            case 15:
                                hora_Ecuador_string = "08";
                                break;
                            case 14:
                                hora_Ecuador_string = "07";
                                break;
                            case 13:
                                hora_Ecuador_string = "06";
                                break;
                            case 12:
                                hora_Ecuador_string = "05";
                                break;
                            case 11:
                                hora_Ecuador_string = "04";
                                break;
                            case 10:
                                hora_Ecuador_string = "03";
                                break;
                            case 9:
                                hora_Ecuador_string = "02";
                                break;
                            case 8:
                                hora_Ecuador_string = "01";
                                break ;
                            case 7:
                                hora_Ecuador_string = "00";
                                break;
                            case 6:
                                hora_Ecuador_string = "23";
                                break;
                            case 5:
                                hora_Ecuador_string = "22";
                                break;
                            case 4:
                                hora_Ecuador_string = "21";
                                break;
                            case 3:
                                hora_Ecuador_string = "20";
                                break;
                            case 2:
                                hora_Ecuador_string = "19";
                                break;
                            case 1:
                                hora_Ecuador_string = "18";
                                break;
                            case 0:
                                hora_Ecuador_string = "17";
                                break;
                            default:
                                break;
                        }
                        //si la hora de Ecuador es menor a 7, se debe bajar un d??a al dia que envio lleida
                        if (hora_Ecuador<7)
                        {
                            dia_Ecuador = dia_Ecuador-1;

                            //si dia_Ecuador es igaul a cero significa que debemos obtener el d??a anterior ya que no existe
                            //el d??a cero en el calendario y restar un mes
                            if (dia_Ecuador==0)
                            {
                                dia_Ecuador = dia_anterior;
                                mes_Ecuador = mes_Ecuador-1;
                            }

                        }
                        fecha_Ecuador = dia_Ecuador + "/" +mes_Ecuador + "/" + linea.substring(11,15)+ " " + hora_Ecuador_string + ":" + linea.substring(21,23)+ ":" + linea.substring(23,25);
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
                        //obtengo el primer correo
                        newStr = mail_to.split(" ",2);
                        mail_to=newStr[0];
                        if (mail_to.equals(correoAnterior)){
                            mail_to="fotomultas-noreply@epmtsd.gob.ec";
                        }
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
                        informationString.append("gstatus_aux: "+gstatus_aux);
                        informationString.append("\n");
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
                        Correo p=new Correo(mail_id,mail_date,fecha_Ecuador,mail_type,file_doc_model,file_uid,unidades_certificadas,mail_from,mail_to,direccion_CC,gstatus,gstatus_aux,mail_subj,add_id,add_displaydate,add_uid);

                        informationString.append("\n");
                        contador++;
                        row=sheet.createRow(contador);

                        row.createCell(numeroCelda).setCellValue(mail_id);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_date);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(fecha_Ecuador);
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
                        correoAnterior=mail_to;

                        list.add(p);
                        //seteo de variables
                        mail_date=unidades_certificadas=mail_type=fecha_Ecuador=file_doc_model=file_uid=mail_from=mail_to=gstatus=gstatus_aux=mail_subj=add_id=add_uid=add_displaydate="";
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
                mensaje="El excel est?? en el escritorio";
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

    public int obtenerContador() {
        return contador;
    }

    public Workbook reportePorFechasSofy(String fechaInicial, String fechaFinal) {
        String fecha=null;
        list.clear();
        link="https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user="+user+"&password="+password+"&mail_date_min="+fechaInicial+"070000&mail_date_max="+fechaFinal+"070000";
        contador = 0;
        int numeroCelda = 0;
        book=null;
        // Creamos el libro de trabajo de Excel formato OOXML
        book = new XSSFWorkbook();



        // La hoja donde pondremos los datos
        Sheet sheet = book.createSheet("Detalle de Reporte");
        //creamos una fila
        Row row = sheet.createRow(contador);
        //row.createCell(1).setCellValue("Reporte mes a??o");
        /////
        /*
        XSSFCell cell1= (XSSFCell) row.createCell(1);
        XSSFCellStyle headerStyle = (XSSFCellStyle) book.createCellStyle();
        Font headerFont = book.createFont();
        headerFont.setColor(IndexedColors.BLUE.getIndex());
        headerStyle.setFont(headerFont);
        cell1.setCellStyle(headerStyle);
        cell1.setCellValue("REPORTE "+fechaInicial.substring(6,8));
        */
        ///creo estilo

        CellStyle style = book.createCellStyle();
        Font font = book.createFont();
        font.setUnderline(Font.SYMBOL_CHARSET);
        font.setColor(IndexedColors.BLUE.getIndex());
        font.setBold(true);//Seteando fuente negrita al encabezado del archivo excel
        //font.setFontHeightInPoints((short) 14);//tama??o de letra
        //font.setItalic(true);//cursiva
        //font.setStrikeout(true);//tachado
        style.setFillForegroundColor((short)100);
        style.setFont(font);

        XSSFCell cell1= (XSSFCell) row.createCell(1);
        cell1.setCellStyle(style);
        cell1.setCellValue("REPORTE "+fechaInicial.substring(6,8));
        //
        contador++;
        row = sheet.createRow(contador);
        row.createCell(1).setCellValue("Detalle de correos certificados");
        contador++;
        row = sheet.createRow(contador);
        //numeroCelda++;
        cell1= (XSSFCell) row.createCell(0);
        cell1.setCellStyle(style);
        cell1.setCellValue("Fecha");
        cell1= (XSSFCell) row.createCell(1);
        cell1.setCellStyle(style);
        cell1.setCellValue("Correo Remitente");
        cell1= (XSSFCell) row.createCell(2);
        cell1.setCellStyle(style);
        cell1.setCellValue("Correo Destinatario");
        cell1= (XSSFCell) row.createCell(3);
        cell1.setCellStyle(style);
        cell1.setCellValue("Correo Certificado");
        cell1= (XSSFCell) row.createCell(4);
        cell1.setCellStyle(style);
        cell1.setCellValue("Correo Recibido");
        cell1= (XSSFCell) row.createCell(5);
        cell1.setCellStyle(style);
        cell1.setCellValue("Correo Leido");


        //esta variable servir?? para almacenar el n??mero del d??a anterior, sirve para poder cambiar
        //la fecha que nos entrega lleida a la fecha de Ecuador ya que el servidor de lleida est??
        //adelantado en relaci??n de  nosotros por 7 horas
        int dia_anterior = 1;
        String[] newStr = null;
        String correoAnterior = "";

        try {

            URL url= new URL(link);
            HttpsURLConnection conn=(HttpsURLConnection)url.openConnection();
            conn.setRequestMethod("GET");
            conn.connect();

            int responseCod= conn.getResponseCode();
            if (responseCod!=200) {
                throw new RuntimeException("ocurrio un error: "+responseCod);
            }else  {
                add_id="";
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
                        //int tamano=linea.length();
                        //int fin=tamano-12;

                        //dia_Ecuador, permite capturar el valor del d??a del correo de lleida
                        //hora_Ecuador, permite capturar el valor de la hora del correo de lleida
                        //mes_Ecuador, permite capturar el valor del mes del correo de lleida
                        int dia_Ecuador, hora_Ecuador,mes_Ecuador;
                        // hora_Ecuador_string, permite capturar la hora del correo en formato string,
                        //se setea en 01 solo por evitar el try catch
                        String hora_Ecuador_string="01";
                        mes_Ecuador = Integer.parseInt(linea.substring(15,17));
                        dia_Ecuador = Integer.parseInt(linea.substring(17,19));
                        hora_Ecuador = Integer.parseInt(linea.substring(19,21));
                        //se compara si el d??a anterior que se seteo en 1, es menosr que el d??a capturado
                        //desde lleida, as?? podemos setear el d??a de Ecuador. Ejemplo si el correo desde
                        //lleida me llega con fecha 01/03/2022 05h:30m:20s yo debo tener el d??a anterior
                        //se el 30 o 31 y poner fecha de Ecuador como 30/02/2022 22h:30m:20s
                        if (dia_anterior <=dia_Ecuador)
                        {
                            dia_anterior = dia_Ecuador;
                        }
                        //como se que la hora de lleida tiene 7 horas en adelanto, se debe setear la hora de Ecuador
                        switch (hora_Ecuador)
                        {
                            case 23:
                                hora_Ecuador_string = "16";
                                break;
                            case 22:
                                hora_Ecuador_string = "15";
                                break;
                            case 21:
                                hora_Ecuador_string = "14";
                                break;
                            case 20:
                                hora_Ecuador_string = "13";
                                break;
                            case 19:
                                hora_Ecuador_string = "12";
                                break;
                            case 18:
                                hora_Ecuador_string = "11";
                                break;
                            case 17:
                                hora_Ecuador_string = "10";
                                break;
                            case 16:
                                hora_Ecuador_string = "09";
                                break;
                            case 15:
                                hora_Ecuador_string = "08";
                                break;
                            case 14:
                                hora_Ecuador_string = "07";
                                break;
                            case 13:
                                hora_Ecuador_string = "06";
                                break;
                            case 12:
                                hora_Ecuador_string = "05";
                                break;
                            case 11:
                                hora_Ecuador_string = "04";
                                break;
                            case 10:
                                hora_Ecuador_string = "03";
                                break;
                            case 9:
                                hora_Ecuador_string = "02";
                                break;
                            case 8:
                                hora_Ecuador_string = "01";
                                break ;
                            case 7:
                                hora_Ecuador_string = "00";
                                break;
                            case 6:
                                hora_Ecuador_string = "23";
                                break;
                            case 5:
                                hora_Ecuador_string = "22";
                                break;
                            case 4:
                                hora_Ecuador_string = "21";
                                break;
                            case 3:
                                hora_Ecuador_string = "20";
                                break;
                            case 2:
                                hora_Ecuador_string = "19";
                                break;
                            case 1:
                                hora_Ecuador_string = "18";
                                break;
                            case 0:
                                hora_Ecuador_string = "17";
                                break;
                            default:
                                break;
                        }
                        //si la hora de Ecuador es menor a 7, se debe bajar un d??a al dia que envio lleida
                        if (hora_Ecuador<7)
                        {
                            dia_Ecuador = dia_Ecuador-1;

                            //si dia_Ecuador es igaul a cero significa que debemos obtener el d??a anterior ya que no existe
                            //el d??a cero en el calendario y restar un mes
                            if (dia_Ecuador==0)
                            {
                                dia_Ecuador = dia_anterior;
                                mes_Ecuador = mes_Ecuador-1;
                            }

                        }
                        fecha_Ecuador = dia_Ecuador + "/" +mes_Ecuador + "/" + linea.substring(11,15)+ " " + hora_Ecuador_string + ":" + linea.substring(21,23)+ ":" + linea.substring(23,25);
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
                        //obtengo el primer correo
                        newStr = mail_to.split(" ",2);
                        mail_to=newStr[0];
                        if (mail_to.equals(correoAnterior)){
                            mail_to="fotomultas-noreply@epmtsd.gob.ec";
                        }
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
                        informationString.append("gstatus_aux: "+gstatus_aux);
                        informationString.append("\n");
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

                        Correo p=new Correo(mail_id,mail_date,fecha_Ecuador,mail_type,file_doc_model,file_uid,unidades_certificadas,mail_from,mail_to,direccion_CC,gstatus,gstatus_aux,mail_subj,add_id,add_displaydate,add_uid);


                        contador++;
                        row=sheet.createRow(contador);

                        row.createCell(numeroCelda).setCellValue(fecha_Ecuador);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_from);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_to);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue("SI");
                        numeroCelda++;
                        if(file_doc_model.equals("KO")){
                            file_doc_model="NO";
                        }else {
                            file_doc_model="YES";
                        }
                        row.createCell(numeroCelda).setCellValue(file_doc_model);
                        numeroCelda++;
                        if (add_id.equals("")){
                            add_id="NO";
                        }else {
                            add_id="SI";
                        }
                        System.out.println(add_id);

                        row.createCell(numeroCelda).setCellValue(add_id);

                        numeroCelda++;
                        correoAnterior=mail_to;
                        list.add(p);
                        //seteo de variables
                        mail_date=unidades_certificadas=mail_type=fecha_Ecuador=file_doc_model=file_uid=mail_from=mail_to=gstatus=gstatus_aux=mail_subj=add_id=add_uid=add_displaydate="";
                        //nueva fila
                        numeroCelda=0;
                    }

                }
                contador++;
                row=sheet.createRow(contador);
                row.createCell(0).setCellValue("Por SODIG S.A.: ");
                contador=contador+5;
                row=sheet.createRow(contador);
                row.createCell(0).setCellValue("Generado por:");
                row.createCell(1).setCellValue("Sof??a Soria");
                row.createCell(2).setCellValue("Revisado por: Marco Logacho ");
                row.createCell(4).setCellValue("Aprobado por: Ivan Robalino");
                contador++;
                row=sheet.createRow(contador);
                row.createCell(3).setCellValue("Aprob?? por el cliente:");
                contador=contador+2;
                row=sheet.createRow(contador);
                row.createCell(3).setCellValue("Ing. Nataly Correa Vallejo");
                contador++;
                row=sheet.createRow(contador);
                row.createCell(3).setCellValue("Procurador Com??n");
                contador++;
                row=sheet.createRow(contador);
                row.createCell(3).setCellValue("Procurador Com??n");
                contador++;
                row=sheet.createRow(contador);
                row.createCell(3).setCellValue("CONSORCIO VIAL SD");

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

    public Workbook reportePorFechasMarco(String fechaInicial, String fechaFinal) {
        String fecha=null;
        list.clear();
        link="https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user="+user+"&password="+password+"&mail_date_min="+fechaInicial+"070000&mail_date_max="+fechaFinal+"070000";
        contador = 0;
        int numeroCelda = 0;
        book=null;
        // Creamos el libro de trabajo de Excel formato OOXML
        book = new XSSFWorkbook();
        // La hoja donde pondremos los datos
        Sheet sheet = book.createSheet("Detalle de Reporte");
        //creamos una fila
        Row row = sheet.createRow(contador);
        row = sheet.createRow(contador);
        //numeroCelda++;
        row.createCell(0).setCellValue("Fecha");
        row.createCell(1).setCellValue("Correo Remitente");
        row.createCell(2).setCellValue("Correo Destinatario");
        row.createCell(3).setCellValue("Correo Certificado");
        row.createCell(4).setCellValue("Correo Recibido");
        row.createCell(5).setCellValue("Correo Leido");


        //esta variable servir?? para almacenar el n??mero del d??a anterior, sirve para poder cambiar
        //la fecha que nos entrega lleida a la fecha de Ecuador ya que el servidor de lleida est??
        //adelantado en relaci??n de  nosotros por 7 horas
        int dia_anterior = 1;
        String[] newStr = null;
        String correoAnterior = "";

        try {

            URL url= new URL(link);
            HttpsURLConnection conn=(HttpsURLConnection)url.openConnection();
            conn.setRequestMethod("GET");
            conn.connect();

            int responseCod= conn.getResponseCode();
            if (responseCod!=200) {
                throw new RuntimeException("ocurrio un error: "+responseCod);
            }else  {
                add_id="";
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
                        //int tamano=linea.length();
                        //int fin=tamano-12;

                        //dia_Ecuador, permite capturar el valor del d??a del correo de lleida
                        //hora_Ecuador, permite capturar el valor de la hora del correo de lleida
                        //mes_Ecuador, permite capturar el valor del mes del correo de lleida
                        int dia_Ecuador, hora_Ecuador,mes_Ecuador;
                        // hora_Ecuador_string, permite capturar la hora del correo en formato string,
                        //se setea en 01 solo por evitar el try catch
                        String hora_Ecuador_string="01";
                        mes_Ecuador = Integer.parseInt(linea.substring(15,17));
                        dia_Ecuador = Integer.parseInt(linea.substring(17,19));
                        hora_Ecuador = Integer.parseInt(linea.substring(19,21));
                        //se compara si el d??a anterior que se seteo en 1, es menosr que el d??a capturado
                        //desde lleida, as?? podemos setear el d??a de Ecuador. Ejemplo si el correo desde
                        //lleida me llega con fecha 01/03/2022 05h:30m:20s yo debo tener el d??a anterior
                        //se el 30 o 31 y poner fecha de Ecuador como 30/02/2022 22h:30m:20s
                        if (dia_anterior <=dia_Ecuador)
                        {
                            dia_anterior = dia_Ecuador;
                        }
                        //como se que la hora de lleida tiene 7 horas en adelanto, se debe setear la hora de Ecuador
                        switch (hora_Ecuador)
                        {
                            case 23:
                                hora_Ecuador_string = "16";
                                break;
                            case 22:
                                hora_Ecuador_string = "15";
                                break;
                            case 21:
                                hora_Ecuador_string = "14";
                                break;
                            case 20:
                                hora_Ecuador_string = "13";
                                break;
                            case 19:
                                hora_Ecuador_string = "12";
                                break;
                            case 18:
                                hora_Ecuador_string = "11";
                                break;
                            case 17:
                                hora_Ecuador_string = "10";
                                break;
                            case 16:
                                hora_Ecuador_string = "09";
                                break;
                            case 15:
                                hora_Ecuador_string = "08";
                                break;
                            case 14:
                                hora_Ecuador_string = "07";
                                break;
                            case 13:
                                hora_Ecuador_string = "06";
                                break;
                            case 12:
                                hora_Ecuador_string = "05";
                                break;
                            case 11:
                                hora_Ecuador_string = "04";
                                break;
                            case 10:
                                hora_Ecuador_string = "03";
                                break;
                            case 9:
                                hora_Ecuador_string = "02";
                                break;
                            case 8:
                                hora_Ecuador_string = "01";
                                break ;
                            case 7:
                                hora_Ecuador_string = "00";
                                break;
                            case 6:
                                hora_Ecuador_string = "23";
                                break;
                            case 5:
                                hora_Ecuador_string = "22";
                                break;
                            case 4:
                                hora_Ecuador_string = "21";
                                break;
                            case 3:
                                hora_Ecuador_string = "20";
                                break;
                            case 2:
                                hora_Ecuador_string = "19";
                                break;
                            case 1:
                                hora_Ecuador_string = "18";
                                break;
                            case 0:
                                hora_Ecuador_string = "17";
                                break;
                            default:
                                break;
                        }
                        //si la hora de Ecuador es menor a 7, se debe bajar un d??a al dia que envio lleida
                        if (hora_Ecuador<7)
                        {
                            dia_Ecuador = dia_Ecuador-1;

                            //si dia_Ecuador es igaul a cero significa que debemos obtener el d??a anterior ya que no existe
                            //el d??a cero en el calendario y restar un mes
                            if (dia_Ecuador==0)
                            {
                                dia_Ecuador = dia_anterior;
                                mes_Ecuador = mes_Ecuador-1;
                            }

                        }
                        fecha_Ecuador = dia_Ecuador + "/" +mes_Ecuador + "/" + linea.substring(11,15)+ " " + hora_Ecuador_string + ":" + linea.substring(21,23)+ ":" + linea.substring(23,25);
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
                        //obtengo el primer correo
                        newStr = mail_to.split(" ",2);
                        mail_to=newStr[0];
                        if (mail_to.equals(correoAnterior)){
                            mail_to="fotomultas-noreply@epmtsd.gob.ec";
                        }
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
                        informationString.append("gstatus_aux: "+gstatus_aux);
                        informationString.append("\n");
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

                        Correo p=new Correo(mail_id,mail_date,fecha_Ecuador,mail_type,file_doc_model,file_uid,unidades_certificadas,mail_from,mail_to,direccion_CC,gstatus,gstatus_aux,mail_subj,add_id,add_displaydate,add_uid);


                        contador++;
                        row=sheet.createRow(contador);

                        row.createCell(numeroCelda).setCellValue(fecha_Ecuador);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_from);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue(mail_to);
                        numeroCelda++;
                        row.createCell(numeroCelda).setCellValue("SI");
                        numeroCelda++;
                        if(file_doc_model.equals("KO")){
                            file_doc_model="NO";
                        }else {
                            file_doc_model="YES";
                        }
                        row.createCell(numeroCelda).setCellValue(file_doc_model);
                        numeroCelda++;
                        if (add_id.equals("")){
                            add_id="NO";
                        }else {
                            add_id="SI";
                        }
                        System.out.println(add_id);

                        row.createCell(numeroCelda).setCellValue(add_id);

                        numeroCelda++;
                        correoAnterior=mail_to;
                        list.add(p);
                        //seteo de variables
                        mail_date=unidades_certificadas=mail_type=fecha_Ecuador=file_doc_model=file_uid=mail_from=mail_to=gstatus=gstatus_aux=mail_subj=add_id=add_uid=add_displaydate="";
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
}