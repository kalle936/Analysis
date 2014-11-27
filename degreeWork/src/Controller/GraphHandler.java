/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Controller;

import Model.DatasetCreator;
import java.io.IOException;
import jxl.read.biff.BiffException;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.time.TimeSeriesCollection;

/**
 * Controller class to make it unable for the view layer to interact with the
 * model layer
 *
 * @author Kalgus
 */
public class GraphHandler {

    public static TimeSeriesCollection getTimeSeries() throws IOException, BiffException, InterruptedException {
        return DatasetCreator.getTimeSeries();
    }

    public static DefaultCategoryDataset getRoomDataset() throws IOException, BiffException {
        return DatasetCreator.getRoomDataset();
    }

    public static TimeSeriesCollection getRoomTimeDataset(String choice) throws IOException, BiffException {
        return DatasetCreator.getRoomTimeDataset(choice);
    }

}
