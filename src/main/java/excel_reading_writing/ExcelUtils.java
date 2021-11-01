package excel_reading_writing;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtils {
    private int rowCount;
    private List<Data> dataList = new ArrayList<Data>();
    private List<Data> modList = new ArrayList<Data>();


    public void getRow() {
        try {
            String excelPath = "/Users/qayumi/Desktop/All_Codes/FullWebAppStore/excel/Data/edu_data.xlsx";
//        String localPath="./Data.edu_data.xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println(rowCount);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void getCell() {
        try {
            String excelPath = "/Users/qayumi/Desktop/All_Codes/FullWebAppStore/excel/Data/edu_data.xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            for (int i = 0; i < rowCount; i++) {
                dataList.add(new Data(
                        String.valueOf(sheet.getRow(i).getCell(0).toString()),
                        String.valueOf(sheet.getRow(i).getCell(1).toString()),
                        String.valueOf(sheet.getRow(i).getCell(2).toString()),
                        String.valueOf(sheet.getRow(i).getCell(3).toString()),
                        String.valueOf(sheet.getRow(i).getCell(4).toString()),
                        String.valueOf(sheet.getRow(i).getCell(5).toString()),
                        String.valueOf(sheet.getRow(i).getCell(6).toString()),
                        String.valueOf(sheet.getRow(i).getCell(7).toString()),
                        String.valueOf(sheet.getRow(i).getCell(8).toString()),
                        String.valueOf(sheet.getRow(i).getCell(9).toString()),
                        String.valueOf(sheet.getRow(i).getCell(10).toString()),
                        String.valueOf(sheet.getRow(i).getCell(11).toString()),
                        String.valueOf(sheet.getRow(i).getCell(12).toString()),
                        String.valueOf(sheet.getRow(i).getCell(13).toString()),
                        String.valueOf(sheet.getRow(i).getCell(14).toString()),
                        String.valueOf(sheet.getRow(i).getCell(15).toString())
                ));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.getRow();
        excelUtils.getCell();

    }

    public void standardSapma() {
        double sum = 0.0;
        double standardDeviation = 0.0;
        double mean = 0.0;
        double res = 0.0;
        double sq = 0.0;

        int n = modList.size();

        System.out.println("Elements are:");
        for (Data data : modList) {
            sum = sum + Double.parseDouble(data.getDiscussion());
        }

        mean = sum / (n);

        for (Data data : modList) {

            standardDeviation
                    = standardDeviation + Math.pow((Double.parseDouble(data.getDiscussion()) - mean), 2);

        }

        sq = standardDeviation / n;
        res = Math.sqrt(sq);
        System.out.println("Standard Sapma " + res);
        System.out.println("Ortalama " + mean);
        System.out.println("Veri setimizin sayisi " + n);
        System.out.println("Merkezden 2 standard sapma degeri " + (mean + 2 * res));

        for (Data data : modList) {
            if (Double.parseDouble(data.getAnnounce()) >= (mean + 2 * res)) {
                System.out.println("removed for Announce column " + data);
//                modList.remove(data);
            }
        }
    }

}
