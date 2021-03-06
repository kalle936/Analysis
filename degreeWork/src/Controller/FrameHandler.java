/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Controller;

import Model.DatasetCreator;
import java.io.IOException;
import java.util.List;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

/**
 * Controller class to make it unable for the view layer to interact with the
 * model layer
 *
 * @author Kalgus
 */
public class FrameHandler {

    public static List showWarnings() throws IOException, WriteException, BiffException {
        return DatasetCreator.showWarnings();
    }

    public static List showPersonalAccess(String name) throws IOException, BiffException {
        return DatasetCreator.showPersonalAccess(name);
    }

}
