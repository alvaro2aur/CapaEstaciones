/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ManejoArchivos;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author alvaro
 */
public class XLSX extends Archivos{
    private XSSFWorkbook libro;
    private XSSFSheet hoja;
    
    
    
    
    public XLSX(String nombre, int nrodehoja) throws FileNotFoundException, IOException{
        super(nombre);
        FileInputStream fis = new FileInputStream(super.archivo);
        libro = new XSSFWorkbook (fis);
        hoja = libro.getSheetAt(nrodehoja);
    }
    
    public int obtenerNroDeHojas(){
        return libro.getNumberOfSheets();
    }
    
    public String obtenerNombreHoja(){
        return hoja.getSheetName();
    }
    
    public ArrayList listaColumna(int columna){
        boolean flag = true;
        ArrayList lista = new ArrayList();
        Iterator<Row> rowIterator = hoja.iterator();
        while (rowIterator.hasNext()) {
            Row fila = rowIterator.next();
            if(flag){
                fila = rowIterator.next();
            }
            Cell celda = fila.getCell(columna);
            if(!lista.contains(celda.getStringCellValue())){
                lista.add(celda.getStringCellValue());
            }
            flag = false;
        }
        return lista;
    }
    
    public ArrayList listaColumna(String nombre, int columna){
        boolean flag = true;
        ArrayList lista = new ArrayList();
        Iterator<Row> rowIterator = hoja.iterator();
        while (rowIterator.hasNext()) {
            Row fila = rowIterator.next();
            if(flag){
                fila = rowIterator.next();
            }
            Cell celda = fila.getCell(columna);
            if(!lista.contains(celda.getStringCellValue())){
                lista.add(celda.getStringCellValue());
            }
            flag = false;
        }
        return lista;
    }
    
    
    
    public String obtenerCelda(int filaa, int columna){
        String valor = "hola";
        Row fila = hoja.getRow(filaa);
        Cell celda = fila.getCell(columna);
        switch (celda.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                valor = celda.getStringCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                valor= Double.toString(celda.getNumericCellValue());
                break;
            default :
        }
        return valor;
        
    }
    
    
    
    
    public int nroDeFilas(){
        boolean flag = true;
        ArrayList lista = new ArrayList();
        Iterator<Row> rowIterator = hoja.iterator();
        int i=0;
        
        while (rowIterator.hasNext()) {
            Row fila = rowIterator.next();
            i++;
        }
        return i;
        
    }
        
}
    

