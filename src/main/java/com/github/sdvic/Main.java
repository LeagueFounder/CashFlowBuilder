package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200823
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import javax.swing.*;

public class Main implements Runnable
{
    private final String version = "200823";
    public PandLReader pandLReader;
    private BudgetReader budgetReader = new BudgetReader();
    private BudgetWriter budgetWriter = new BudgetWriter();
    private CashItemAggregator cashItemAggregator = new CashItemAggregator();
    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }
    public void run()
    {
        System.out.println("Version " + version + "\nCopyright 2020 Vic Wintriss\n");
        int targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month"));
        String followOnAnswer = JOptionPane.showInputDialog("Initial Budget Run? (Yes or Return for No");
        pandLReader = new PandLReader();
        budgetReader = new BudgetReader();
        pandLReader.readPandL(targetMonth);//Read QuickBooks Excel P&L into HashMap
        budgetReader.readBudget(targetMonth, followOnAnswer);//Read Budget Excel into HashMap
        cashItemAggregator.extractBudgetItemsFromPandLMap(pandLReader.getPandlHashMap(), budgetReader.getBudgetHashMap());
        cashItemAggregator.extractBudgetItemsFromBudgetMap(pandLReader.getPandlHashMap(), budgetReader.getBudgetHashMap(), targetMonth);
        cashItemAggregator.computeCombinedCashBudgetSheetEntries();
        cashItemAggregator.updateBudgetWorkbook(budgetReader.getBudgetWorkBook(), targetMonth);
        budgetWriter.writeBudget(budgetReader.getBudgetWorkBook());
    }
}
