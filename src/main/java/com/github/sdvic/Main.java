package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L
 * rev 0.3 March 29, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

public class Main
{
    public static File PandLfile;
    public static FileInputStream fis;
    public static XSSFWorkbook pandlWorkbook;
    public static XSSFWorkbook sarah5yearWorkbook;
    public static HashMap<Cell, Cell> accountList = new HashMap<Cell, Cell>();
    public static String k;
    public static int v;
    public static int cashContributions;
    public static Sheet pandlSheet;
    public static Sheet sarah5yearSheet;
    public static int pandlLastRow;
    public static int sarah5yearLastRow;
    public static Row pandlRow;
    public static Row sarah5yearRow;
    public static Cell cell;
    public static Cell valueCell;
    private static String key;
    private static Cell keyCell;

    public static void main(String[] args) throws IOException
    {
        DataFormatter formatter = new DataFormatter();
        try
        {
            pandlWorkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/VicMini/Desktop/LeaguePandL.xlsx")));
            pandlSheet = pandlWorkbook.getSheetAt(0);
            pandlLastRow = pandlSheet.getLastRowNum();
            sarah5yearWorkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx")));

            sarah5yearSheet = sarah5yearWorkbook.getSheetAt(0);
            sarah5yearLastRow = sarah5yearSheet.getLastRowNum();
        }
        catch (Exception e)
        {
            System.out.println("Can't read /Users/VicMini/Desktop/LeaguePandL.xlsx");
        }
        System.out.println("========================== P & L =====================");
        for (int i = 0; i <= pandlLastRow; i++)
        {
            pandlRow = pandlSheet.getRow(i);
            if (pandlRow != null)
            {
                keyCell = pandlRow.getCell(0);
                valueCell = pandlRow.getCell(1);
                System.out.print(keyCell + "    " + valueCell);
            }
        }
        for (int i = 0; i <= sarah5yearLastRow; i++)
        {

            if (sarah5yearSheet.getRow(i).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK)
            {
                sarah5yearRow = sarah5yearSheet.getRow(i);
                keyCell = sarah5yearRow.getCell(0);
                valueCell = sarah5yearRow.getCell(6);
                System.out.println(keyCell + "    " + valueCell);
                Cell testCell = sarah5yearSheet.getRow(i).getCell(12);
                testCell.setCellValue("Row " + i);
            }
        }
        try
        {
            FileOutputStream out = new FileOutputStream(new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx"));
            sarah5yearWorkbook.write(out);
            out.close();
        }
        catch(Exception e)
        {
            System.out.println("Can't write to /Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
        }
    }
}
