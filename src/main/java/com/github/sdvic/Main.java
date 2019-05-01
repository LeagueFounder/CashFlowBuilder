package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190501d
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;

public class Main implements Runnable
{
    public  File pandLfile;
    public  FileInputStream plfis;
    public  XSSFWorkbook pandlWorkbook;//P&L
    private HashMap<String, Integer> chartOfAccountsMap;

    public  File sarah5YearFile;
    public  FileInputStream s5yrfis;
    public  XSSFWorkbook sarah5yearWorkbook;//5Year
    ExcelReader excelReader;
    ExcelWriter excelWriter;
    private String version = "version 190501d";
    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        System.out.println(version);
        try
        {
            pandLfile = new File("/Users/VicMini/Desktop/PandL2018.xlsx");
            plfis = new FileInputStream(pandLfile);
            pandlWorkbook = new XSSFWorkbook(plfis);

            sarah5YearFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
            s5yrfis = new FileInputStream(sarah5YearFile);
            sarah5yearWorkbook = new XSSFWorkbook(s5yrfis);
            excelReader = new ExcelReader(pandlWorkbook, sarah5yearWorkbook, version);
            chartOfAccountsMap =  excelReader.getChartOfAccountsMap();
            new CashFlowItemAggregator(sarah5yearWorkbook, chartOfAccountsMap, version);
            excelWriter = new ExcelWriter(sarah5yearWorkbook, sarah5YearFile);
            excelWriter.write5YearPlan();

        }
        catch (Exception e)
        {
            System.out.println("Exception in Main.run()" + e);
        }
    }
}


