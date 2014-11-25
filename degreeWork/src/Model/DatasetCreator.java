package Model;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;

public class DatasetCreator {

    private DatasetCreator() throws BiffException, WriteException, IOException {
    }
    private static List uniqueDateList = new ArrayList();

    /**
     * Method that creates a list of formatted strings containing all warnings
     * that exist in the excel file.
     *
     * @return List of warnings.
     * @throws IOException
     * @throws WriteException
     * @throws jxl.read.biff.BiffException
     */
    public static List showWarnings() throws IOException, WriteException, BiffException {
        WorkbookSettings ws = new WorkbookSettings();
        ws.setEncoding("Cp1252");
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"), ws);
        Sheet sheet = workbook.getSheet(0);
        DateCell dateCell;
        Cell nameCell;
        Cell affectedCell;
        Cell temperatureCell;

        List warningDataset = new ArrayList();
        int rowsToCheck = sheet.getRows();

        for (int i = 1; i < rowsToCheck; i++) {
            nameCell = sheet.getCell(1, i);
            if (nameCell.getContents().contains("Varning")) {
                dateCell = (DateCell) sheet.getCell(0, i);
                affectedCell = sheet.getCell(2, i);
                temperatureCell = sheet.getCell(5, i);

                if (!warningDataset.contains(dateCell.getDate() + " | " + nameCell.getContents() + " | " + affectedCell.getContents())) {
                    warningDataset.add(dateCell.getDate() + " | " + nameCell.getContents() + " | " + affectedCell.getContents() + " | " + temperatureCell.getContents());
                }
            }

        }
        workbook.close();
        return warningDataset;
    }

    /**
     * method that creates a list of a certain persons accesses. It ignores
     * duplicates (these might exist in the excel file).
     *
     * @param name The name of the person whose accesses you want to find.
     * @return List of the specific persons accesses
     * @throws IOException
     * @throws BiffException
     */
    public static List showPersonalAccess(String name) throws IOException, BiffException {
        WorkbookSettings ws = new WorkbookSettings();
        ws.setEncoding("Cp1252");
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"), ws);
        Sheet sheet = workbook.getSheet(0);

        DateCell dateCell;
        Cell nameCell;
        Cell eventCell;
        Cell affectedCell;
        List personalDataset = new ArrayList();
        int rowsToCheck = sheet.getRows();
        SimpleDateFormat dt = new SimpleDateFormat("yyyy-MM-dd hh:mm");

        for (int i = 1; i < rowsToCheck; i++) {
            nameCell = sheet.getCell(3, i);
            if (nameCell.getContents().contains(name)) {
                dateCell = (DateCell) sheet.getCell(0, i);
                eventCell = sheet.getCell(1, i);
                affectedCell = sheet.getCell(2, i);
                if (!personalDataset.contains(dt.format(dateCell.getDate()) + " | " + eventCell.getContents() + " | " + affectedCell.getContents())) {
                    personalDataset.add(dt.format(dateCell.getDate()) + " | " + eventCell.getContents() + " | " + affectedCell.getContents());
                }
            }

        }

        workbook.close();
        return personalDataset;

    }

    /**
     * Method that picks out all the date of all accesses made in the excel file
     * and counts how many there are. they are countet with respect to the date
     * that they were registered in the excel file.
     *
     * @return
     * @throws IOException
     * @throws BiffException
     * @throws InterruptedException
     */
    public static TimeSeriesCollection getTimeSeries() throws IOException, BiffException, InterruptedException {
        WorkbookSettings ws = new WorkbookSettings();
        ws.setEncoding("Cp1252");
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"), ws);
        Sheet sheet = workbook.getSheet(0);
        TimeSeries series = new TimeSeries("time series", Day.class);
        TimeSeriesCollection dataset = new TimeSeriesCollection();

        int rowsToCheck = sheet.getRows();
        int day;
        int month;
        int year;
        int number;
        Calendar cal = Calendar.getInstance();
        List daySeen = new ArrayList();
        Day dayRead;
        for (int i = 1; i < rowsToCheck; i++) {

            DateCell dateCell = (DateCell) sheet.getCell(0, i);
            Date date = dateCell.getDate();
            cal.setTime(date);
            day = cal.get(Calendar.DAY_OF_MONTH);
            month = cal.get(Calendar.MONTH) + 1;
            year = cal.get(Calendar.YEAR);
            dayRead = new Day(day, month, year);
            Cell actionCell = sheet.getCell(1, i);
            String action = actionCell.getContents();

            if (action.equalsIgnoreCase("Dörr Upplåst") || action.equals("Tvångsöppnad")) {
                if (isUnique(date)) {
                    daySeen.add(dayRead);
                    number = countNumberEqual(daySeen, dayRead);
                    series.addOrUpdate(dayRead, number);
                }
            }
        }

        dataset.addSeries(series);
        uniqueDateList.clear();
        workbook.close();
        return dataset;
    }

    /**
     * Method that creates the dataset nessesary for building a bar-graph in the
     * view layer.
     *
     * @return returns a DefaultCategoryDataset that is needed to create a bar
     * graph.
     * @throws IOException
     * @throws BiffException
     */
    @SuppressWarnings("empty-statement")
    public static DefaultCategoryDataset getRoomDataset() throws IOException, BiffException {
        WorkbookSettings ws = new WorkbookSettings();
        ws.setEncoding("Cp1252");
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"), ws);
        Sheet sheet = workbook.getSheet(0);

        int rowsToCheck = sheet.getRows();
        int receptionCount12 = 0;
        int lagerCount12 = 0;
        int plan5Count12 = 0;
        int nyckelrumCount12 = 0;
        int receptionCount13 = 0;
        int lagerCount13 = 0;
        int plan5Count13 = 0;
        int nyckelrumCount13 = 0;
        int receptionCount14 = 0;
        int lagerCount14 = 0;
        int plan5Count14 = 0;
        int nyckelrumCount14 = 0;
        int day;
        List timestamps = new ArrayList();
        Calendar cal = Calendar.getInstance();
        DefaultCategoryDataset objDataset = new DefaultCategoryDataset();

        for (int i = 1; i < rowsToCheck; i++) {
            Cell roomCell = sheet.getCell(2, i);
            DateCell dateCell = (DateCell) sheet.getCell(0, i);
            Date date = dateCell.getDate();
            cal.setTime(date);
            day = cal.get(Calendar.DAY_OF_MONTH);
            String roomName = roomCell.getContents();

            if (isUnique(date)) {

                if (roomName.contains("7001") && day == 12) { //Reception
                    receptionCount12++;
                } else if (roomName.contains("11001") && day == 12) { //Entrédörr lager
                    lagerCount12++;
                } else if (roomName.contains("3002") && day == 12) { //trapphus
                    plan5Count12++;
                } else if (roomName.contains("14002") && day == 12) { //nyckelrum
                    nyckelrumCount12++;
                } else if (roomName.contains("7001") && day == 13) { //Reception
                    receptionCount13++;
                } else if (roomName.contains("11001") && day == 13) { //Entrédörr lager
                    lagerCount13++;
                } else if (roomName.contains("3002") && day == 13) { //trapphus
                    plan5Count13++;
                } else if (roomName.contains("14002") && day == 13) { //nyckelrum
                    nyckelrumCount13++;
                } else if (roomName.contains("7001") && day == 14) { //Reception
                    receptionCount14++;
                } else if (roomName.contains("11001") && day == 14) { //Entrédörr lager
                    lagerCount14++;
                } else if (roomName.contains("3002") && day == 14) { //trapphus
                    plan5Count14++;
                } else if (roomName.contains("14002") && day == 14) { //nyckelrum
                    nyckelrumCount14++;
                }
            } else {
                timestamps.add(date);
            }
        }
        objDataset.setValue(receptionCount12, "Reception", "12/10");
        objDataset.setValue(receptionCount13, "Reception", "13/10");
        objDataset.setValue(receptionCount14, "Reception", "14/10");
        objDataset.setValue(lagerCount12, "Entrédörr lager", "12/10");
        objDataset.setValue(lagerCount13, "Entrédörr lager", "13/10");
        objDataset.setValue(lagerCount14, "Entrédörr lager", "14/10");
        objDataset.setValue(plan5Count12, "Plan 5", "12/10");
        objDataset.setValue(plan5Count13, "Plan 5", "13/10");
        objDataset.setValue(plan5Count14, "Plan 5", "14/10");
        objDataset.setValue(nyckelrumCount12, "Nyckelrum", "12/10");
        objDataset.setValue(nyckelrumCount13, "Nyckelrum", "13/10");
        objDataset.setValue(nyckelrumCount14, "Nyckelrum", "14/10");
        workbook.close();
        uniqueDateList.clear();
        return objDataset;
    }

    public static TimeSeriesCollection getRoomTimeDataset(String choice) throws IOException, BiffException {
        WorkbookSettings ws = new WorkbookSettings();
        ws.setEncoding("Cp1252");
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"), ws);
        Sheet sheet = workbook.getSheet(0);
        TimeSeries series = new TimeSeries("time series", Day.class);
        TimeSeriesCollection dataset = new TimeSeriesCollection();

        int rowsToCheck = sheet.getRows();
        int day;
        int month;
        int year;
        int number;
        Calendar cal = Calendar.getInstance();
        List daySeen = new ArrayList();
        Day dayRead;
        for (int i = 1; i < rowsToCheck; i++) {

            DateCell dateCell = (DateCell) sheet.getCell(0, i);
            Date date = dateCell.getDate();
            cal.setTime(date);
            day = cal.get(Calendar.DAY_OF_MONTH);
            month = cal.get(Calendar.MONTH) + 1;
            year = cal.get(Calendar.YEAR);
            dayRead = new Day(day, month, year);
            Cell doorCell = sheet.getCell(2, i);
            String door = doorCell.getContents();

            if (isUnique(date)) {
                if (door.contains(choice)) {
                    daySeen.add(dayRead);
                    number = countNumberEqual(daySeen, dayRead);
                    series.addOrUpdate(dayRead, number);
                }
            }
        }

        dataset.addSeries(series);
        workbook.close();
        uniqueDateList.clear();
        return dataset;
    }

    /**
     * Counts how many ocurrances of a certain Day already exists in a list.
     *
     * @param itemList
     * @param itemToCheck
     * @return
     */
    private static int countNumberEqual(List itemList, Day itemToCheck) {
        int count = 0;
        for (Object i : itemList) {
            if (i.equals(itemToCheck)) {
                count++;
            }
        }
        return count;
    }

    private static boolean isUnique(Date date) {
        if (uniqueDateList.contains(date)) {
            return false;
        } else {
            uniqueDateList.add(date);
            return true;
        }
    }
}
