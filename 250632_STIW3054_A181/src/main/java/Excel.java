
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;


public class Excel 
{

    private static final List<arrangement> info = new ArrayList();
    private static final String URL = "https://ms.wikipedia.org/wiki/Malaysia";

    public static void Web_Table() 
    {
        try 
        {
            System.out.println("It will take few seconds");
            System.out.println("Connecting"  + URL + ".....");

            Document source = Jsoup.connect(URL).get();
            Element table = source.select("Table").get(5);
            Elements rows = table.select("Tr");

            for (Element row : rows) 
            {

                Elements data1 = row.select("Th");
                Elements data2 = row.select("Td");

                String column1 = data1.text();
                String column2 = data2.text();

                info.add(new arrangement(column1, column2));
            }

            System.out.println("Data successfully collected from table Trivia .");
            System.out.println();

        } 
        catch (IOException e) 
        {
            System.out.println("ERROR : Connection Failed " + URL);
        }
    }

    public static void Excel() 
    {

        if (info.isEmpty()) 
        {
            System.out.println("ERROR : Terminated..... No data to write.");
            System.exit(0); //without this, the program will write empty excel file
        }

        String excelFile = "Output File.xlsx";

        System.out.println("Writing  " + excelFile + "...");

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Malaysia");
        try 
        {
            for (int i = 0; i < info.size(); i++) 
            {
                XSSFRow row = sheet.createRow(i);

                XSSFCell cell1 = row.createCell(0);
                cell1.setCellValue(info.get(i).getHeader());
                XSSFCell cell2 = row.createCell(1);
                cell2.setCellValue(info.get(i).getData());
                
            }

            FileOutputStream outputFile = new FileOutputStream(excelFile);
            workbook.write(outputFile);
            outputFile.flush();
            outputFile.close();
            System.out.println(excelFile + " Writting Completed .");
        } 
        catch (IOException e) 
        {
            System.out.println("ERROR : Writting Failed !");
        }
    }

    public static void main(String[] args) 
    {
        Web_Table();
        Excel();
    }    
}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Admin
 */
