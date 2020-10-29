package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201029B
// * copyright 2020 Vic Wintriss
//*******************************************************************************************
import javax.swing.*;
public class Main implements Runnable
{
    public static void main(String[] args)
    {
        SwingUtilities.invokeLater(new Main());
    }
    public void run()
    {
        String version = "201029";
        System.out.println("Version " + version + "\nCopyright 2020 Vic Wintriss\n");
        int targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month"));
        String followOnAnswer = JOptionPane.showInputDialog("Initial Budget Run? (Yes or Return for No");
        PandLReader pandLReader = new PandLReader();
        BudgetReader budgetReader = new BudgetReader();
        BudgetWriter budgetWriter = new BudgetWriter();
        CashItemAggregator cashItemAggregator = new CashItemAggregator();
        pandLReader.readPandL(targetMonth);//Read QuickBooks Excel P&L into pandlHashMap
        budgetReader.readBudget(targetMonth, followOnAnswer);//Read Original/Updated Excel Budget into budgetHashMap
        cashItemAggregator.computeCombinedCashBudgetSheetEntries(pandLReader.getPandlMap(), budgetReader.getBudgetMap(), budgetReader.getBudgetWorkBook(), targetMonth);
        cashItemAggregator.updateBudgetWorkbook(budgetReader.getBudgetWorkBook(), targetMonth);
        budgetWriter.writeBudget(budgetReader.getBudgetWorkBook());
    }
}
