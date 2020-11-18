package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201117A
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
        String version = "201110";
        System.out.println("Version " + version + "\nCopyright 2020 Vic Wintriss\n");
        int targetMonth = Integer.parseInt(JOptionPane.showInputDialog("Please enter QuickBooks P and L input file (int)month"));
        String followOnAnswer = JOptionPane.showInputDialog("Initial Budget Run? (Yes or Return for No");
//        targetMonth = 7;
//        followOnAnswer = "Yes";
        PandLReader pandLReader = new PandLReader(targetMonth);
        BudgetReader budgetReader = new BudgetReader(targetMonth, followOnAnswer);
        BudgetWriter budgetWriter = new BudgetWriter();
        CashItemAggregator cashItemAggregator = new CashItemAggregator(budgetReader.getBudgetMap(), pandLReader.getPandLMap(), targetMonth);
        cashItemAggregator.computeDonations(); //Consolidating P&L entries to Budget entries
        cashItemAggregator.computeTuition();
        cashItemAggregator.computeMiscIncome();
        cashItemAggregator.computeTotalIncome();
        cashItemAggregator.computeSalaries();
        cashItemAggregator.computeContractServices();
        cashItemAggregator.computeRent();
        cashItemAggregator.computeOperatons();
        cashItemAggregator.computeMiscExpense();
        cashItemAggregator.computeTotalExpenses();
        cashItemAggregator.computeProfit();
        cashItemAggregator.computeStudents();
        cashItemAggregator.reconcile();
        cashItemAggregator.updateBudgetWorkbook(budgetReader.getBudgetWorkBook(), targetMonth);
        budgetWriter.writeBudget(budgetReader.getBudgetWorkBook());
    }
}
