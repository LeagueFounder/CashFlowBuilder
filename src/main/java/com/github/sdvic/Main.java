package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.4 April 17, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/*******************************************************************************
 *   Extract P&L data from QuickBooks, analyze and insert into a cash flow chart
 *   ver 1.3 April 17, 2019
 *******************************************************************************/
public class Main implements Runnable
{
    public  XSSFWorkbook pandlWorkbook;//P&L
    public  XSSFWorkbook sarah5yearLocalWorkbook;//5Year
    public  Sheet pandlSheet;//P&L
    public  Sheet sarah5yearLocalSheet;//5Year
    public  File pandLfile;
    public  File sarah5YearRemoteFile;
    public  FileInputStream plfis;
    public  FileInputStream s5yrfis;
    public  FileOutputStream plfos;
    public  FileOutputStream s5yrfos;

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
            sarah5YearRemoteFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
            s5yrfis = new FileInputStream(sarah5YearRemoteFile);
            sarah5yearLocalWorkbook = (XSSFWorkbook) WorkbookFactory.create(s5yrfis);
            sarah5yearLocalSheet = sarah5yearLocalWorkbook.getSheetAt(0);
        }
        catch (Exception e)
        {
            System.out.println("Exception in setup()" + e);
        }
        new ExcelReader(pandlSheet, sarah5yearLocalSheet);
        new ExcelWriter(sarah5yearLocalSheet);
    }
}


