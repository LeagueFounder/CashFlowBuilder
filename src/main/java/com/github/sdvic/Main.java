package com.github.sdvic;
/********************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200625A
 * copyright 2020 Vic Wintriss
 /*******************************************************************************************/

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.io.*;
import java.util.HashMap;

import static org.apache.poi.ss.usermodel.WorkbookFactory.create;

public class Main implements Runnable
{
    private final String version = "version 200627B";
    private final File pandlFile = new File("/Users/VicMini/Desktop/The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx");
    private final File budgetFile = new File("/Users/VicMini/Desktop/LeagueBudget2020.xlsx");
    private FileInputStream pandlFIS;
    private FileInputStream budgetFIS;
    private FileOutputStream budgetFOS;
    private XSSFWorkbook pandlWorkbook;
    private XSSFWorkbook budgetWorkBook;
    public ExcelReader excelReader;
    private BudgetWriter budgetWriter;
    private HashMap<String, Integer> pandlMap = new HashMap<>();
    private CashItemAggregator aggregator = new CashItemAggregator(budgetWorkBook, pandlMap, version);

    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        System.out.println(version);
        JOptionPane.showInputDialog("P and L input file month?  Only enter (int)month.");
        System.out.println("processing " + pandlFile);
        System.out.println("processing " + budgetFile);
        try
        {
            pandlFIS = new FileInputStream(pandlFile);
            budgetFIS = new FileInputStream(budgetFile);
            budgetFOS = new FileOutputStream(budgetFile);
            pandlWorkbook = new XSSFWorkbook(pandlFIS);
            budgetWorkBook = new XSSFWorkbook(pandlFIS);
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        excelReader = new ExcelReader(pandlWorkbook, budgetWorkBook, version);
        //budgetWriter = new BudgetWriter(budgetWorkBook, budgetFOS);
        System.out.println(budgetWorkBook.getSheetAt(0));

    }
}
