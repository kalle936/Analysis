/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package View;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

/**
 *
 * @author Kalgus
 */
public class StartUp {
    /**
     * main method that creates the GUI and makes it visible to the user. Also modifies the look and feel 
     * of the graphical user interface.
     *
     * @param args the command line arguments
     * @throws jxl.write.WriteException
     * @throws jxl.read.biff.BiffException
     */
    public static void main(String args[]) throws WriteException, BiffException {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ApplicationView.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            try {
                new customDialog().setVisible(true);
                new ApplicationView().setVisible(true);
                
            } catch (BiffException | WriteException | IOException ex) {
                Logger.getLogger(ApplicationView.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
    }  
}
