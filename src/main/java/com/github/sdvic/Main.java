package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L
 * rev 1.0 April 3, 2019
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
    public static File pandLfile = new File("/Users/VicMini/Desktop/PandL2018.xlsx");
    public static File sarahFiveYearFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
    public static FileInputStream plfis;
    public static FileInputStream s5yrfis;
    public static Sheet sheet;

    public static void main(String[] args) throws IOException
    {
        System.out.println("Cash Flow Generator ver 1.0  4/3/19");
        System.out.println("========================== Reading P & L =====================");
        try
        {
            plfis = new FileInputStream(pandLfile);
            pandlWorkbook = (XSSFWorkbook) WorkbookFactory.create(plfis);
            System.out.println("P and L Workbook has " + pandlWorkbook.getNumberOfSheets() + " Sheet");
            pandlSheet = pandlWorkbook.getSheetAt(0);//P&L
            System.out.println("pandL row 0, column 0 => " + pandlSheet.getRow(0).getCell(0));
        }
        catch(Exception e)
        {
            System.out.println("Can't read pandL input file");
        }
        sheet = pandlSheet;
        int totalTuition = findPandLentry("Total 47201 Tuition  Fees");
        int totalWorkshop = findPandLentry("Total 47202 Workshop Fees");
        int totalContributions = findPandLentry("43450 Individ, Business Contributions");
        System.out.println("Total Cash Income " + (totalContributions + totalTuition + totalWorkshop));

        System.out.println("========================== Reading 5-YearPlan =====================");
        try
        {
            s5yrfis = new FileInputStream(sarahFiveYearFile);
            sarah5yearWorkbook = (XSSFWorkbook)WorkbookFactory.create(s5yrfis);
            System.out.println("Sarah Workbook has " + sarah5yearWorkbook.getNumberOfSheets() + " Sheet");
            sarah5yearSheet = sarah5yearWorkbook.getSheetAt(0);
            System.out.println("Sarah5yr row 0, column 0 => " + sarah5yearSheet.getRow(0).getCell(0));
        }
        catch (Exception e)
        {
            System.out.println("Can't read Sarah input file");
        }
        plfis.close();
        s5yrfis.close();
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
