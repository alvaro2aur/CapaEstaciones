/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ManejoArchivos;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;

/**
 *
 * @author alvaro
 */
public class TXT extends Archivos {
    
    FileWriter a;
    BufferedWriter b;
    PrintWriter txt;
        
    public TXT(String nombre) throws IOException{
        super(nombre);
        a = new FileWriter(super.archivo);
        b = new BufferedWriter(a);
        txt = new PrintWriter(b);
    }
    
    public void escribirLinea(String texto){
        txt.println(texto);
        
    }
    
    public void cerrarArchivo(){
        txt.close();
        
    }
    
    
}
