package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.6 April 27, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Main implements Runnable
{
    public  File pandLfile;
    public  FileInputStream plfis;
    public  XSSFWorkbook pandlWorkbook;//P&L
    public  Sheet pandlSheet;//P&L

    public  File sarah5YearFile;
    public  FileInputStream s5yrfis;
    public  XSSFWorkbook sarah5yearWorkbook;//5Year
    public  Sheet sarah5yearSheet;//5Year

    public static void main(String[] args)
    {
        System.out.println("Cash Flow Generator ver 1.2 4/15/19");
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        try
        {
            pandLfile = new File("/Users/VicMini/Desktop/PandL2018.xlsx");
            plfis = new FileInputStream(pandLfile);
            pandlWorkbook = (XSSFWorkbook) WorkbookFactory.create(plfis);
            pandlSheet = pandlWorkbook.getSheetAt(0);//P&L

            sarah5YearFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
            s5yrfis = new FileInputStream(sarah5YearFile);
            sarah5yearWorkbook = (XSSFWorkbook) WorkbookFactory.create(s5yrfis);
            sarah5yearSheet = sarah5yearWorkbook.getSheetAt(0);
            new ExcelReader(pandlSheet, sarah5yearSheet, sarah5yearWorkbook);
        }
        catch (Exception e)
        {
            System.out.println("Exception in setup()" + e);
        }

    }
}


