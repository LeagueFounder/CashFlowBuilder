package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L
 * rev 0.4 March 31, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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
    public static int totalTuition;
    private static Cell pandLnameCell;

    public static void main(String[] args) throws IOException
    {
        System.out.println("Cash Flow Generator ver 0.4  3/31/19");
        try
        {
            pandlWorkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/VicMini/Desktop/PandL2018.xlsx")));//P&L
            pandlSheet = pandlWorkbook.getSheetAt(0);//P&L
            pandlLastRow = pandlSheet.getLastRowNum();//P&L
            sarah5yearWorkbook = new XSSFWorkbook(new FileInputStream(new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx")));//5Year
            sarah5yearSheet = sarah5yearWorkbook.getSheetAt(0);//5Year
            sarah5yearLastRow = sarah5yearSheet.getLastRowNum();//5Year
        }
        catch (Exception e)
        {
            System.out.println("Can't read input file(s)");
        }
        System.out.println("========================== P & L =====================");
        for (int i = 0; i <= pandlLastRow; i++)
        {
            pandlRow = pandlSheet.getRow(i);
            if (pandlRow != null)
            {
                pandLnameCell = pandlRow.getCell(0);
                pandLvalueCell = pandlRow.getCell(1);
                if (pandLnameCell.getStringCellValue().contains("Total 47201 Tuition"))
                {
                    totalTuition = (int) pandLvalueCell.getNumericCellValue();
                    System.out.println("bingo...value => " + totalTuition);
                }
            }
        }
        System.out.println("========================== 5-YearPlan =====================");
        Row newDataRow = sarah5yearSheet.getRow(5);
        Cell newDataCelll = newDataRow.getCell(6);
        newDataCelll.setCellValue(totalTuition);

        System.out.println("Adding updated info (" + totalTuition + ") to the  " +  " Column, ");
        try
        {
            FileOutputStream sarah5YearFileOutputStream = new FileOutputStream("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
            sarah5yearWorkbook.write(sarah5YearFileOutputStream);
            sarah5YearFileOutputStream.close();

        }
        catch (Exception e)
        {
            System.out.println("Can't write to /Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
        }
    }
}

