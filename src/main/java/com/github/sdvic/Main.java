package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200809
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileOutputStream;
import java.util.HashMap;
public class Main implements Runnable
{
    private final String version = "200809";
    public PandLReader pandLReader;
    private BudgetReader budgetReader = new BudgetReader();
    public BudgetWriter budgetWriter = new BudgetWriter();
    private CashItemAggregator cashItemAggregator = new CashItemAggregator();
    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }
    public void run()
    {
        System.out.println("Version " + version + "\nCopyright 2020 Vic Wintriss");
        int targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month"));
        pandLReader = new PandLReader();
        budgetReader = new BudgetReader();
        pandLReader.readPandLtoHashMap(targetMonth);
        budgetReader.readBudget();
        cashItemAggregator.aggregateBudget(pandLReader.getPandlHashMap(), budgetReader.getBudgetHashMap(), targetMonth);
        cashItemAggregator.computeLineItems();
        cashItemAggregator.updateBudgetWorkbook(budgetReader.getBudgetWorkBook(), targetMonth);
        budgetWriter.writeBudget(budgetReader.getBudgetWorkBook());
    }
}
