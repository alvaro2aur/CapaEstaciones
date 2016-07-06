/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ManejoArchivos;

import java.io.File;

/**
 *
 * @author alvaro
 */
public class Archivos {
    
    File archivo;
    
    public Archivos(String nombre){
        archivo = new File(nombre);
    }
    
}
