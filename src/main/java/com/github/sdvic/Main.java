package com.github.sdvic;
/********************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200702A
 * copyright 2020 Vic Wintriss
 /*******************************************************************************************/
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.io.*;
import java.util.HashMap;

public class Main implements Runnable
{
    private final String version = "version 200703D";
    private final File pandlFile = new File("/Users/VicMini/Desktop/The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx");
    private final File budgetFile = new File("/Users/VicMini/Desktop/XXXVicBudget2020.xlsx");
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

    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        System.out.println(version);
        targetMonth = Integer.parseInt(JOptionPane.showInputDialog("P and L input file month?  Only enter (int)month."));
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
        System.out.println("(7) created budgetFOS => " + budgetFOS);
        excelReader = new ExcelReader(pandlWorkbook, budgetWorkBook, version);
        pandlMap = excelReader.getPandLMap();
        aggregator = new CashItemAggregator(budgetWorkBook, pandlMap, targetMonth);
        budgetWorkBook.getSheetAt(0).getRow(0).getCell(0).setCellValue(version);
        try
        {
            budgetWorkBook.write(budgetFOS);
        }
        catch (IOException e)
        {
            System.out.println("!!!!!!!!!!! (8)Write budgetWorkBook error");
            e.printStackTrace();
        }
        System.out.println("(8)wrote out to budget workbook");
//        budgetWriter = new BudgetWriter(budgetWorkBook, budgetFOS);
    }
}
