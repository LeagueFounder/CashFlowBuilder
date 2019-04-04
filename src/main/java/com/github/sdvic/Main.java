package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L
 * rev 0.8 April 2, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main
{
    public static XSSFWorkbook pandlWorkbook;//P&L
    public static XSSFWorkbook sarah5yearWorkbook;//5Year
    public static Sheet pandlSheet;//P&L
    public static Sheet sarah5yearSheet;//5Year
    public static Row pandlRow;//P&L
    public static Row sarah5yearRow;//5Year
    public static int pandlLastRow;//P&L
    public static int sarah5yearLastRow;//5Year
    public static Cell pandLvalueCell;
    public static Cell pandLnameCell;
    public static Cell sarah5yearValueCell;
    public static int pAndLvalue;
    public static File pandLfile = new File("/Users/VicMini/Desktop/PandL2018.xlsx");
    public static File sarahFiveYearFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
    public static OutputStream sarah5YearOutputStream;
    public static InputStream sarah5YearInputStream;
    public static InputStream pandLInputStream;
    public static Workbook pandLworkbook;
    public static DataFormatter formatter;
    public static Sheet sheet;

    public static void main(String[] args) throws IOException
    {
        System.out.println("Cash Flow Generator ver 0.8  4/2/19");
        try
        {
            pandLworkbook = WorkbookFactory.create(new FileInputStream("/Users/VicMini/Desktop/PandL2018.xlsx"));
            System.out.println("P and L Workbook has " + pandLworkbook.getNumberOfSheets() + " Sheets : ");
            pandLInputStream = new FileInputStream(pandLfile);
            sarah5YearInputStream = new FileInputStream(sarahFiveYearFile);
            sarah5YearOutputStream = new FileOutputStream(sarahFiveYearFile);
            pandlWorkbook = new XSSFWorkbook(pandLInputStream);//P&L
            //sarah5yearWorkbook = new XSSFWorkbook(sarah5YearInputStream);//5Year
            pandlSheet = pandlWorkbook.getSheetAt(0);//P&L
            //sarah5yearSheet = sarah5yearWorkbook.getSheetAt(0);//5Year
        }
        catch (Exception e)
        {
            System.out.println("Can't read input file(s)");
        }
        System.out.println("========================== P & L =====================");
        sheet = pandLworkbook.getSheetAt(0);
        int totalTuition = findPandLentry("Total 47201 Tuition  Fees");
        int totalWorkshop = findPandLentry("Total 47202 Workshop Fees");
        int totalContributions = findPandLentry("43450 Individ, Business Contributions");
        System.out.println(totalContributions + totalTuition + totalWorkshop);
        System.out.println("========================== 5-YearPlan =====================");
        try
        {
            //sarah5yearWorkbook.write(sarah5YearOutputStream);
            sarah5YearOutputStream.close();
            pandLInputStream.close();
        }
        catch (Exception e)
        {
            System.out.println("Can't write File");
        }
    }

    public static int findPandLentry(String name)
    {
        for (Row row : sheet)
        {
            for (Cell cell : row)
            {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING)
                {
                    if (cell.getStringCellValue().contains(name))
                    {
                        int cellValue = (int) row.getCell(1).getNumericCellValue();
                        return cellValue;
                    }
                }
            }
        }
        return 4;
    }
}
