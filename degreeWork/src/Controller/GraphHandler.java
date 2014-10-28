/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Controller;

import Model.BusinessLogic;
import java.io.IOException;
import jxl.read.biff.BiffException;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.time.TimeSeriesCollection;

/**
 *
 * @author Kalgus
 */
public class GraphHandler {

    public static TimeSeriesCollection getTimeSeries() throws IOException, BiffException {
        return BusinessLogic.getTimeSeries();
    }

    public static DefaultCategoryDataset getRoomDataset() throws IOException, BiffException {
        return BusinessLogic.getRoomDataset();
    }

}
