package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200709A
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileOutputStream;
import java.util.HashMap;

public class Main implements Runnable
{
    private final String version = "200709A";
    private FileOutputStream budgetOutputFOS;
    private XSSFWorkbook pandlWorkbook;
    private XSSFSheet budgetSheet;
    public PandLReader pandLtoHashMapReader = new PandLReader();
    public BudgetWriter budgetWriter = new BudgetWriter();
    private HashMap<String, Integer> pandlHashMap = new HashMap<>();
    private CashItemAggregator aggregator = new CashItemAggregator();
    private int targetMonth;
    private BudgetReader budgetReader = new BudgetReader();

    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }

    public void run()
    {
        targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month."));
        budgetReader.readBudget();
        pandLtoHashMapReader.readPandLtoHashMap();
        aggregator.aggregateBudget(budgetReader.getBudgetWorkBook(),pandLtoHashMapReader.getPandlHashMap(), targetMonth);
        budgetWriter.writeBudget(aggregator.getBudgetWorkbook());
    }
}
