package com.github.sdvic;
/****************************************************************************************
 * * Application to extract Cash Flow data from Quick Books P&L
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static org.apache.poi.ss.usermodel.Cell.*;

public class Main
{
    public static File file;
    public static FileInputStream fis;
    public static XSSFWorkbook workbook;
    public static Map<String, Integer> accountList = new HashMap<String, Integer>();
    public static String v;
    public static int k;
    public static int cashContributions;


    public static void main(String[] args) throws IOException
    {
        DataFormatter formatter = new DataFormatter();
        try
        {
            file = new File("/Users/VicMini/Desktop/LeaguePandL.xlsx");
            fis = new FileInputStream(file);
            workbook = new XSSFWorkbook(fis);
        }
        catch (Exception e)
        {
            System.out.println("Can't read /Users/VicMini/Desktop/LeaguePandL.xlsx");
        }
        for (Sheet sheet : workbook)
        {
            System.out.println(" There is (are) " + workbook.getNumberOfSheets() + " sheet(s) in PandL.xlsx");
            for (Row row : sheet)
            {
                for (Cell cell : row)
                {

                    switch (cell.getCellType())
                    {
                        case CELL_TYPE_BOOLEAN:
                            break;
                        case CELL_TYPE_NUMERIC:
                            k = (int) cell.getNumericCellValue();//key = dollars
                            break;
                        case CELL_TYPE_STRING:
                            try
                            {
                                v = cell.getStringCellValue();//val = text
                                accountList.put(v,k);
                            }
                            catch (NumberFormatException ex)
                            {
                                System.out.println("Number Format Exception");
                            }
                            break;
                        case CELL_TYPE_BLANK:
                            break;
                        case CELL_TYPE_FORMULA:
                            System.out.println("formula");
                            break;
                        default:
                            System.out.println("default error...no cell type discovered");
                    }
                }
            }
        }
        fis.close();
        for (Map.Entry<String, Integer> entry : accountList.entrySet())
        {
            String key = entry.getKey();
            Integer value = entry.getValue();
            if (key.contains("47200"))
            {
                cashContributions = value;
                System.out.println(value);
            }
        }
       // accountList.forEach((k, v) -> System.out.println(k + "k" + v + "v"));
    }
}

