package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200717
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileOutputStream;
import java.util.HashMap;
public class Main implements Runnable
{
    private final String version = "200717";
    private FileOutputStream budgetOutputFOS;
    private XSSFWorkbook pandlWorkbook;
    private XSSFSheet budgetSheet;
    public PandLReader pandLReader = new PandLReader();
    private BudgetReader budgetReader = new BudgetReader();
    public BudgetWriter budgetWriter = new BudgetWriter();
    private HashMap<String, Integer> pandlHashMap = new HashMap<>();
    private CashItemAggregator cashItemAggregator = new CashItemAggregator();
    private int targetMonth;

    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }
    public void run()
    {
        targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month."));
        budgetReader.readBudget();
        pandLReader.readPandLtoHashMap();
        cashItemAggregator.aggregateBudget(budgetReader.getBudgetWorkBook(), pandLReader.getPandlHashMap(), targetMonth);
        budgetWriter.writeBudget(cashItemAggregator.getBudgetWorkbook());

    }
}
