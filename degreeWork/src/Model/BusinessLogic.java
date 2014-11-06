package Model;

import java.awt.BorderLayout;
import java.awt.Container;
import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import javax.swing.BorderFactory;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.border.TitledBorder;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;

public class BusinessLogic {

    private BusinessLogic() throws BiffException, WriteException, IOException {
    }

    /**
     * Method that copies the information in the main excel file into the
     * temporary holding place where changes can be made.
     *
     * @return
     * @throws IOException
     * @throws WriteException
     * @throws jxl.read.biff.BiffException
     */
    public static List showWarnings() throws IOException, WriteException, BiffException {
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"));
        Sheet sheet = workbook.getSheet(0);

        DateCell dateCell;
        Cell nameCell;
        Cell affectedCell;
        Cell temperatureCell;

        List warningList = new ArrayList();
        int rowsToCheck = sheet.getRows();

        for (int i = 1; i < rowsToCheck; i++) {
            nameCell = sheet.getCell(1, i);
            if (nameCell.getContents().contains("Varning")) {
                dateCell = (DateCell) sheet.getCell(0, i);
                affectedCell = sheet.getCell(2, i);
                temperatureCell = sheet.getCell(5, i);

                if (!warningList.contains(dateCell.getDate() + " | " + nameCell.getContents() + " | " + affectedCell.getContents())) {
                    warningList.add(dateCell.getDate() + " | " + nameCell.getContents() + " | " + affectedCell.getContents() + " | " + temperatureCell.getContents());
                }
            }

        }
        workbook.close();
        return warningList;
    }

    public static List showPersonalAccess(String name) throws IOException, BiffException {
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"));
        Sheet sheet = workbook.getSheet(0);

        DateCell dateCell;
        Cell nameCell;
        Cell eventCell;
        Cell affectedCell;
        List personalList = new ArrayList();
        int rowsToCheck = sheet.getRows();
        double progress = 0;
        SimpleDateFormat dt = new SimpleDateFormat("yyyy-MM-dd hh:mm");

        for (int i = 1; i < rowsToCheck; i++) {
            nameCell = sheet.getCell(3, i);
            if (nameCell.getContents().contains(name)) {
                dateCell = (DateCell) sheet.getCell(0, i);
                eventCell = sheet.getCell(1, i);
                affectedCell = sheet.getCell(2, i);
                if (!personalList.contains(dt.format(dateCell.getDate()) + " | " + eventCell.getContents() + " | " + affectedCell.getContents())) {
                    personalList.add(dt.format(dateCell.getDate()) + " | " + eventCell.getContents() + " | " + affectedCell.getContents());
                }
            }

        }

        workbook.close();
        return personalList;

    }

    public static TimeSeriesCollection getTimeSeries() throws IOException, BiffException, InterruptedException {

        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"));
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
            daySeen.add(dayRead);
            number = countNumberEqual(daySeen, dayRead);
            series.addOrUpdate(dayRead, number);
        }

        dataset.addSeries(series);
        workbook.close();
        return dataset;
    }

    @SuppressWarnings("empty-statement")
    public static DefaultCategoryDataset getRoomDataset() throws IOException, BiffException {
        Workbook workbook = Workbook.getWorkbook(new File("C:\\Users\\Kalgus\\Documents\\Events Macces 1 vecka.xls"));
        Sheet sheet = workbook.getSheet(0);

        int rowsToCheck = sheet.getRows();
        int reception12Count = 0;
        int lager12Count = 0;
        int trapp12Count = 0;
        int nyckelrum12Count = 0;
        int reception13Count = 0;
        int lager13Count = 0;
        int trapp13Count = 0;
        int nyckelrum13Count = 0;
        int reception14Count = 0;
        int lager14Count = 0;
        int trapp14Count = 0;
        int nyckelrum14Count = 0;
        int day;
        Calendar cal = Calendar.getInstance();
        DefaultCategoryDataset objDataset = new DefaultCategoryDataset();

        for (int i = 1; i < rowsToCheck; i++) {
            Cell roomCell = sheet.getCell(2, i);
            DateCell dateCell = (DateCell) sheet.getCell(0, i);
            Date date = dateCell.getDate();
            cal.setTime(date);
            day = cal.get(Calendar.DAY_OF_MONTH);
            String roomName = roomCell.getContents();
            if (roomName.contains("7001") && day == 12) { //Reception
                reception12Count++;
            }
            if (roomName.contains("11001") && day == 12) { //Entrédörr lager
                lager12Count++;
            }
            if (roomName.contains("3002") && day == 12) { //trapphus
                trapp12Count++;
            }
            if (roomName.contains("14002") && day == 12) { //nyckelrum
                nyckelrum12Count++;
            }
            if (roomName.contains("7001") && day == 13) { //Reception
                reception13Count++;
            }
            if (roomName.contains("11001") && day == 13) { //Entrédörr lager
                lager13Count++;
            }
            if (roomName.contains("3002") && day == 13) { //trapphus
                trapp13Count++;
            }
            if (roomName.contains("14002") && day == 13) { //nyckelrum
                nyckelrum13Count++;
            }
            if (roomName.contains("7001") && day == 14) { //Reception
                reception14Count++;
            }
            if (roomName.contains("11001") && day == 14) { //Entrédörr lager
                lager14Count++;
            }
            if (roomName.contains("3002") && day == 14) { //trapphus
                trapp14Count++;
            }
            if (roomName.contains("14002") && day == 14) { //nyckelrum
                nyckelrum14Count++;
            }
        }
        objDataset.setValue(reception12Count, "12/10", "Reception");
        objDataset.setValue(reception13Count, "13/10", "Reception");
        objDataset.setValue(reception14Count, "14/10", "Reception");
        objDataset.setValue(lager12Count, "12/10", "Entrédörr lager");
        objDataset.setValue(lager13Count, "13/10", "Entrédörr lager");
        objDataset.setValue(lager14Count, "14/10", "Entrédörr lager");
        objDataset.setValue(trapp12Count, "12/10", "Trapphus");
        objDataset.setValue(trapp13Count, "13/10", "Trapphus");
        objDataset.setValue(trapp14Count, "14/10", "Trapphus");
        objDataset.setValue(nyckelrum12Count, "12/10", "Nyckelrum");
        objDataset.setValue(nyckelrum13Count, "13/10", "Nyckelrum");
        objDataset.setValue(nyckelrum14Count, "14/10", "Nyckelrum");
        workbook.close();
        return objDataset;
    }

    private static int countNumberEqual(List itemList, Day itemToCheck) {
        int count = 0;
        for (Object i : itemList) {
            if (i.equals(itemToCheck)) {
                count++;
            }
        }
        return count;
    }

    private static int countAccess(List personSeen, String name) {
        int count = 0;
        for (Object i : personSeen) {
            if (i.equals(name)) {
                count++;
            }
        }
        return count;
    }

}
