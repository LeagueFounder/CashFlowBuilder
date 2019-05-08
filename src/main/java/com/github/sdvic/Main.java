package com.github.sdvic;
/********************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190508
 * copyright 2019 Vic Wintriss
 /*******************************************************************************************/
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.nio.file.FileSystemException;
import java.util.HashMap;

public class Main implements Runnable
{
    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        String version = "version 190508a";
        System.out.println(version);
        String pandlYear = JOptionPane.showInputDialog("P and L input file year?  Only enter year...no other text...no suffix...");
        int yearColumnIndex = (Integer.parseInt(pandlYear) % 2000) - 12;
        System.out.println("processing /Users/VicMini/Desktop/PandL" + pandlYear);
        try
        {
            File pandLfile = new File("/Users/VicMini/Desktop/PandL" + pandlYear + ".xlsx");
            FileInputStream plfis = new FileInputStream(pandLfile);
            XSSFWorkbook pandlWorkbook = new XSSFWorkbook(plfis);
            File sarah5YearFile = new File("/Users/VicMini/Desktop/SarahFiveYearPlan.xlsx");
            FileInputStream s5yrfis = new FileInputStream(sarah5YearFile);
            XSSFWorkbook sarah5yearWorkbook = new XSSFWorkbook(s5yrfis);
            ExcelReader excelReader = new ExcelReader(pandlWorkbook, sarah5yearWorkbook, version);
            HashMap<String, Integer> chartOfAccountsMap = excelReader.getChartOfAccountsMap();
            new CashFlowItemAggregator(sarah5yearWorkbook, chartOfAccountsMap, version, yearColumnIndex);
            ExcelWriter excelWriter = new ExcelWriter(sarah5yearWorkbook, sarah5YearFile);
            excelWriter.write5YearPlan();
        }
        catch (FileSystemException e)
        {
            System.out.println("File exception in Main.run()" + e);
        }
        catch (Exception e)
        {
            System.out.println("exception in Main.run()" + e);
        }
    }
}
