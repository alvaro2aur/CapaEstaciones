/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package capa.estaciones;

import ManejoArchivos.TXT;
import ManejoArchivos.XLSX;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;

/**
 *
 * @author alvaro
 */
public class CapaEstaciones {

    /**
     * @param args the command line arguments
     */
    static ArrayList<String> ctrl = new ArrayList<String>();
    static ArrayList<String> est = new ArrayList<String>();
    static ArrayList<String> nombresest = new ArrayList<String>();
    static Double [][] slice=new Double[40][2];
    
    
    public static void main(String[] args) throws IOException {
        Calendar fecha = new GregorianCalendar();
        
        String nombreinput="Base de datos_Definitiva 10-05-16.xlsx";     //Direccion+nombre+extencion del archivo de entrada
        String nombreoutput="Estaciones "+fecha.get(Calendar.DATE)+"-"+fecha.get(Calendar.MONTH)+"-"+fecha.get(Calendar.YEAR)+".kml";
           
        double longarco;//46465/sdsd/sdsdds/sdsd.xlsx
        double aper;
        
        String opacidad="ff";
        String color2G=opacidad.concat("121212");
        String color3GF1=opacidad.concat("121212");
        String color3GF2=opacidad.concat("121212");
        String color4G=opacidad.concat("121212");
        
        
        
        TXT kml = new TXT(nombreoutput);
        
//        XLSX datos3G = new XLSX(nombreinput,1);
//        XLSX datos4G = new XLSX(nombreinput,2);
//        XLSX config = new XLSX(nombreinput,3);
        
        //Verificacion de tipo de archivo de entrada
        if(!nombreinput.substring(nombreinput.length()-4).contentEquals("xlsx")){
            System.out.println("Por favor verifique el tipo de archivo");
            System.out.println("El archivo debe ser .xlsx");
            System.exit(0);
        }
        
        //Verificacion de la longitud de los colores
        if(color2G.length()!=8 || color3GF1.length()!=8 || color3GF2.length()!=8 || color4G.length()!=8){
            System.out.println("Por favor verifique los codigos de colores");
            System.out.println("Los codigos de colores deben poseer 6 caracteres y la opacidad 2");
            System.exit(0);
        }
        
        encabezado(kml, nombreoutput);
        estilos(kml, color2G, color3GF1, color3GF2, color4G);
        for(int i =0; i<1; i++){
            XLSX datos = new XLSX(nombreinput,i);
            hacerKMZ(datos, kml);
        }
        pie(kml);
        
        
        
    }
    
    public static void hacerKMZ(XLSX datos, TXT kml){
        String celda, estacion, longitud, latitud, azimut, controladora, cellid, cellid2,
                TE, TM, altura, antena, cluster;
        ctrl.clear();
        ctrl = datos.listaColumna(3);//Crea ArrayList de controladoras
        nombresest.clear();
        nombresest = datos.listaColumna(7);
        
        abrirCarpeta(kml, datos.obtenerNombreHoja());
        
        for(int i=0;i<ctrl.size();i++){
            abrirCarpeta(kml, ctrl.get(i));
            est.clear();
            
            abrirCarpeta(kml, "Nombres");
            
            for(int j = 1;j<datos.nroDeFilas();j++){
                estacion = datos.obtenerCelda(j, 7).substring(0, datos.obtenerCelda(j, 7).length()-1);
                controladora = datos.obtenerCelda(j, 3);
                longitud = datos.obtenerCelda(j, 9);
                latitud = datos.obtenerCelda(j, 10);
                 
                if(ctrl.get(i).contains(controladora) && !est.contains(estacion)){
                    est.add(estacion);
                    nombresEstaciones(kml, estacion, longitud, latitud, controladora);
                }                
            }
            
            cerrarCarpeta(kml);//Nombres de estaciones
            
//            imprimirArrayList(est);
//            System.out.println("-------------------");
            
            for(int j=1;j<est.size();j++){
                abrirCarpeta(kml, est.get(j));
                
                for(int k=1;k<datos.nroDeFilas();k++){
                    estacion=datos.obtenerCelda(k, 7).substring(0, datos.obtenerCelda(k, 7).length()-1);
                    
                    if(est.get(j).contentEquals(estacion)){
                        celda=datos.obtenerCelda(k, 6);
                        cluster=datos.obtenerCelda(k, 2);
                        cellid=datos.obtenerCelda(k, 4);
                        cellid2=datos.obtenerCelda(k, 5);
                        celda=datos.obtenerCelda(k, 6);
                        longitud=datos.obtenerCelda(k, 9);
                        latitud=datos.obtenerCelda(k, 10);
                        azimut=datos.obtenerCelda(k, 11);
                        TM=datos.obtenerCelda(k, 12);
                        TE=datos.obtenerCelda(k, 13);
                        altura=datos.obtenerCelda(k, 14);
                        antena=datos.obtenerCelda(k, 20);
                        
                        slice = crearSlice(azimut, longitud, latitud, cellid);
                        imprimirSlice(kml, celda, slice, longitud, latitud, azimut, TE, TM, altura, antena, cluster, cellid);
                        
                        
                        
                        
                        
                    }
                    
                    
                }
                cerrarCarpeta(kml);//Estaciones
                
                
            }
            cerrarCarpeta(kml);//Controladora
            
            
            
        }
        cerrarCarpeta(kml);//Tecnologia
        
        
        
        
        
        
    }
    
    
    
    public static void encabezado(TXT objeto, String nombre){
        objeto.escribirLinea("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        objeto.escribirLinea("<kml xmlns=\"http://earth.google.com/kml/2.2\">");
        objeto.escribirLinea("<Document>");
        objeto.escribirLinea("<name>"+nombre+"</name>");
        objeto.escribirLinea("");
    }
    
    public static void estilos(TXT objeto, String color2G, String color3GF1, String color3GF2, String color4G){
        objeto.escribirLinea("<open>1</open>");
        objeto.escribirLinea("<Style id=\"2G\">");
        objeto.escribirLinea("<LineStyle>");
        objeto.escribirLinea("<color>ff000000</color>");
        objeto.escribirLinea("</LineStyle>");
        objeto.escribirLinea("<PolyStyle>");
        objeto.escribirLinea("<color>"+color2G+"</color>");
        objeto.escribirLinea("</PolyStyle>");
        objeto.escribirLinea("</Style>");
        objeto.escribirLinea("");
        
        objeto.escribirLinea("<Style id=\"3GF1\">");
        objeto.escribirLinea("<LineStyle>");
        objeto.escribirLinea("<color>ff000000</color>");
        objeto.escribirLinea("</LineStyle>");
        objeto.escribirLinea("<PolyStyle>");
        objeto.escribirLinea("<color>"+color3GF1+"</color>");
        objeto.escribirLinea("</PolyStyle>");
        objeto.escribirLinea("</Style>");
        objeto.escribirLinea("");
        
        objeto.escribirLinea("<Style id=\"3GF2\">");
        objeto.escribirLinea("<LineStyle>");
        objeto.escribirLinea("<color>ff000000</color>");
        objeto.escribirLinea("</LineStyle>");
        objeto.escribirLinea("<PolyStyle>");
        objeto.escribirLinea("<color>"+color3GF2+"</color>");
        objeto.escribirLinea("</PolyStyle>");
        objeto.escribirLinea("</Style>");
        objeto.escribirLinea("");
                
        objeto.escribirLinea("<Style id=\"4G\">");
        objeto.escribirLinea("<LineStyle>");
        objeto.escribirLinea("<color>ff000000</color>");
        objeto.escribirLinea("</LineStyle>");
        objeto.escribirLinea("<PolyStyle>");
        objeto.escribirLinea("<color>"+color4G+"</color>");
        objeto.escribirLinea("</PolyStyle>");
        objeto.escribirLinea("</Style>");
        objeto.escribirLinea("");
    }
    
    public static void pie(TXT objeto){
        objeto.escribirLinea("</Document>");
        objeto.escribirLinea("</kml>");
        objeto.cerrarArchivo();
    }
    
    public static void abrirCarpeta(TXT objeto, String nombre){
        objeto.escribirLinea("<Folder>");
        objeto.escribirLinea("<name>"+nombre+"</name>");
    }

    public static void cerrarCarpeta(TXT objeto){
        objeto.escribirLinea("</Folder>");
        objeto.escribirLinea("");
    }
    
    public static void imprimirArrayList(ArrayList array){
        for(int i=0;i<array.size();i++){
            System.out.println(array.get(i));
        }
    }
    
    public static void nombresEstaciones(TXT objeto, String estacion, String longitud, String latitud, String controladora){
        objeto.escribirLinea("<Placemark>");
        objeto.escribirLinea("<Style id=\"textstyle\">");
        objeto.escribirLinea("<IconStyle>");
        objeto.escribirLinea("<Icon>");
        objeto.escribirLinea("</Icon>");
        objeto.escribirLinea("</IconStyle>");
        //objeto.escribirLinea("<LabelStyle>");
        //objeto.escribirLinea("<color>ffffffff</color>");
        //objeto.escribirLinea("</LabelStyle>");
        objeto.escribirLinea("</Style>");
        objeto.escribirLinea("<name>" + estacion + "</name>");
        objeto.escribirLinea("<LookAt>");
        objeto.escribirLinea("<longitude>" + longitud.replace(",", ".") + "</longitude>");
        objeto.escribirLinea("<latitude>" + latitud.replace(",", ".") + "</latitude>");
        objeto.escribirLinea("<altitude>0</altitude>");
        objeto.escribirLinea("<range>10</range>");
        objeto.escribirLinea("<tilt>0</tilt>");
        objeto.escribirLinea("<heading>0</heading>");
        objeto.escribirLinea("</LookAt>");
        objeto.escribirLinea("<styleUrl>#"+controladora+"</styleUrl>");
        objeto.escribirLinea("<Point>");
        objeto.escribirLinea("<coordinates>" + longitud.replace(",", ".") + ", " + latitud.replace(",", ".") + " ,0</coordinates>");
        objeto.escribirLinea("</Point>");
        objeto.escribirLinea("</Placemark>");
        objeto.escribirLinea("");
    }

    public static void imprimirSlice(TXT objeto, String celda, Double matriz[][], String longitud, String latitud, String azimut, String TE, String TM, String altura, String antena, String cluster, String cellid){
        /*
        String estilo="w";
        switch(N){
        case 0:
            estilo="2G";
            break;
        case 1:
            estilo="3G";
            break;
        case 2:
            estilo="4G";
            break;       
        }
        */
        objeto.escribirLinea("<Placemark>");
        objeto.escribirLinea("<name>" + celda + " " + cellid + "</name>");
        //objeto.escribirLinea("<description><![CDATA[LAC = " + lac + "]]></description>");
        objeto.escribirLinea("<description><![CDATA[Longitud = " + longitud.substring(0, longitud.length()-4) +  "<br> Latitud = " + latitud + "<br> Azimut =  " + azimut + "°" + "<br> TE =  " + TE + "°" + "<br> TM =  " + TM + "°" + "<br> Altura =  " + altura + " m" + "<br> Antena =  " + antena + "<br>" + cluster + "]]></description>");
        //objeto.escribirLinea("<LookAt>");
        //objeto.escribirLinea("<longitude>" + longitud +"</longitude>");
        //objeto.escribirLinea("<latitude>" + latitud + "</latitude>");
        //objeto.escribirLinea("<altitude>0</altitude>");
        //objeto.escribirLinea("<range>10</range>");
        //objeto.escribirLinea("<tilt>0</tilt>");
        //objeto.escribirLinea("<heading>0</heading>");
        //objeto.escribirLinea("</LookAt>");
        objeto.escribirLinea("<styleUrl>#2G</styleUrl>");
        objeto.escribirLinea("<Polygon>");
        //objeto.escribirLinea("<altitudeMode>relativeToGround</altitudeMode>");
        objeto.escribirLinea("<altitudeMode>clampToGround</altitudeMode>");
        objeto.escribirLinea("<outerBoundaryIs>");
        objeto.escribirLinea("<LinearRing>");
        objeto.escribirLinea("<coordinates>");
        for(byte i=0;i<=20;i++){
            objeto.escribirLinea(matriz[i][0] + "," + matriz[i][1] + ",0");
        }
        objeto.escribirLinea("</coordinates>");
        objeto.escribirLinea("</LinearRing>");
        objeto.escribirLinea("</outerBoundaryIs>");
        objeto.escribirLinea("</Polygon>");
        objeto.escribirLinea("</Placemark>");
        objeto.escribirLinea("");
    }
    
    public static Double[][] crearSlice(String azimut, String longitud, String latitud, String cellid){
        Double d, lat, longi, alfa, beta, ang, azi, pasos;
        
        d = 0.002;
        Byte aper = 65;
        
        lat = Double.parseDouble(latitud.replace(",", "."));
        longi = Double.parseDouble(longitud.replace(",", "."));
        azi = Double.parseDouble(azimut.replace(",", "."));
                
        slice[0][0] =longi;
        slice[0][1] = lat;
        
        alfa = (azi+(aper/2))*Math.PI/180;
        beta = (azi-(aper/2))*Math.PI/180;
        
        ang=beta;
        pasos=(alfa-beta)/19;
        
        
        
        for (byte i=1; i<=19; i++){
            //System.out.println(i);
            slice[i][0]=longi+(d*Math.cos(ang));
            slice[i][1]=lat+(d*Math.sin(ang));
            ang=ang+pasos;
        }
        
        slice[20][0]=longi;
        slice[20][1]=lat;
        
        return slice;
        
        }
}

