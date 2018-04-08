/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package dataposting;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;



/**
 * 
 * @author Aman Mambetov "asanbay"@mail.ru
 */
public class DataPosting {

static DPoster myPoster;
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        
                
        try {
            LogManager.getLogManager().readConfiguration(
            DataPosting.class.getResourceAsStream("logging.properties"));
        } catch (IOException ex) {
            Logger.getLogger(DataPosting.class.getName()).log(Level.SEVERE, null, ex);
        }

        
        myPoster = new DPoster();
        myPoster.doTask();
        // TODO code application logic here
    }
    
}
