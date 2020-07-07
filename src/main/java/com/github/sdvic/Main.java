package com.github.sdvic;
/********************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projection
 * version 200706A
 * copyright 2020 Vic Wintriss
 /*******************************************************************************************/

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.HashMap;

public class Main implements Runnable
{
    private final String version = "200706A";
    private final File pandlFile = new File("/Users/VicMini/Desktop/The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx");
    private final File budgetFile = new File("/Users/VicMini/Desktop/VicBudget2020.xlsx");
    private FileInputStream pandlFIS;
    private FileInputStream budgetFIS;
    private FileOutputStream budgetFOS;
    private XSSFWorkbook pandlWorkbook;
    private XSSFWorkbook budgetWorkBook;
    public ExcelReader excelReader;
    private BudgetWriter budgetWriter;
    private HashMap<String, Integer> pandlMap = new HashMap<>();
    private CashItemAggregator aggregator;
    private int targetMonth;
    private LocalDateTime now;
    Calendar cals;

    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        cals = Calendar.getInstance();
        targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month."));
        System.out.println("CashFlowBuilder version " + version + " is adding month " + targetMonth + " data from QuickBooks P&L file: " + pandlFile + " to Budget file: " + budgetFile + " at: ");
        System.out.println("(1)============ processing pandlFile => " + pandlFile + " reading month: " + targetMonth);
        System.out.println("(2)========= processing budgetFile => " + budgetFile);
        try
        {
            pandlFIS = new FileInputStream(pandlFile);
        }
        catch (Exception e)
        {
            System.out.println("!!!!!!!!!!! (3)pandlFIS error");
            e.printStackTrace();
        }
        System.out.println("(3)created pandlFIS => " + pandlFIS);
        try
        {
            pandlWorkbook = new XSSFWorkbook(pandlFIS);
        }
        catch (Exception e)
        {
            System.out.println("!!!!!!!!!!! (4)pandlWorkBook error");
            e.printStackTrace();
        }
        System.out.println("(4)created pandlWorkBook => " + pandlWorkbook);
        try
        {
            budgetFIS = new FileInputStream(budgetFile);
        }
        catch (Exception e)
        {
            System.out.println("!!!!!!!!!!! (5)budgetFIS error");
            e.printStackTrace();
        }
        System.out.println("(5) created budgetFIS => " + budgetFIS);
        try
        {
            budgetWorkBook = new XSSFWorkbook(budgetFIS);
        }
        catch (Exception e)
        {
            System.out.println("!!!!!!!!!!! (6)budgetWorkBook error");
            e.printStackTrace();
        }
        System.out.println("(6)Created budgetWorkBook => " + budgetWorkBook);
        try
        {
            budgetFOS = new FileOutputStream(budgetFile);
        }
        catch (Exception e)
        {
            System.out.println("!!!!!!!!!!! (7)budgetFOS error");
            e.printStackTrace();
        }
        System.out.println("(7) created budgetFOS => " + budgetFOS + "\n");
        excelReader = new ExcelReader(pandlWorkbook, budgetWorkBook, version);
        pandlMap = excelReader.getPandLMap();
        aggregator = new CashItemAggregator(budgetWorkBook, pandlMap, targetMonth);
        budgetWorkBook.getSheetAt(0).getRow(0).getCell(0).setCellValue(String.valueOf(cals.getTime()));
        try
        {
            budgetWorkBook.write(budgetFOS);
        }
        catch (IOException e)
        {
            System.out.println("!!!!!!!!!!! (8)Write budgetWorkBook error");
            e.printStackTrace();
        }
        System.out.println("\n(8)wrote out to budget workbook");
//        budgetWriter = new BudgetWriter(budgetWorkBook, budgetFOS);
    }
}
