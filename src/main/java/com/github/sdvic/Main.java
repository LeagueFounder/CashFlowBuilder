package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201120
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
        String version = "201120";
        System.out.println("Version " + version + "\nCopyright 2020 Vic Wintriss\n");
        int targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month"));
        String followOnAnswer = JOptionPane.showInputDialog("Initial Budget Run? (Yes or Return for No");
//        targetMonth = 7;
//        followOnAnswer = "Yes";
        PandLReader pandLReader = new PandLReader(targetMonth);
        BudgetReader budgetReader = new BudgetReader(targetMonth, followOnAnswer);
        BudgetWriter budgetWriter = new BudgetWriter();
        CashItemAggregator cashItemAggregator = new CashItemAggregator(budgetReader.getBudgetMap(), pandLReader.getPandLMap(), targetMonth);
        cashItemAggregator.updateBudgetWorkbook(budgetReader.getBudgetWorkBook(), targetMonth);
        budgetWriter.writeBudget(budgetReader.getBudgetWorkBook());
        cashItemAggregator.reconcile();
    }
}
