package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L
 * rev 1.1 April 8, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/*******************************************************************************
 *   Extract P&L data from QuickBooks, analyze and insert into a cash flow chart
 *   ver 1.2 April 15, 2019
 *******************************************************************************/
public class Main
{
    public static XSSFWorkbook pandlWorkbook;//P&L
    public static XSSFWorkbook sarah5yearLocalWorkbook;//5Year
    public static Sheet pandlSheet;//P&L
    public static Sheet sarah5yearLocalSheet;//5Year
    public static File pandLfile;
    public static File sarahFiveYearRemoteFile;
    public static FileInputStream plfis;
    public static FileInputStream s5yrfis;
    public static FileOutputStream plfos;
    public static FileOutputStream s5yrfos;
    public static Setup setup;

    public static void main(String[] args)
    {
        System.out.println("Cash Flow Generator ver 1.2 4/15/19");
        try
        {
            pandLfile = new File("/Users/VicMini/Desktop/PandL2018.xlsx");
            plfis = new FileInputStream(pandLfile);
            pandlWorkbook = (XSSFWorkbook) WorkbookFactory.create(plfis);
            pandlSheet = pandlWorkbook.getSheetAt(0);//P&L
            sarahFiveYearRemoteFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
            s5yrfis = new FileInputStream(sarahFiveYearRemoteFile);
            sarah5yearLocalWorkbook = (XSSFWorkbook) WorkbookFactory.create(s5yrfis);
            sarah5yearLocalSheet = sarah5yearLocalWorkbook.getSheetAt(0);
        }
        catch (Exception e)
        {
            System.out.println("Exception in setup()" + e);
        }
        setup = new Setup(pandlSheet, sarah5yearLocalSheet);
    }
}

