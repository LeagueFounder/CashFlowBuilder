package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 210115
// * copyright 2020 Vic Wintriss
//*******************************************************************************************
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
public class BudgetWriter
{
    private final File updated2020MasterBudgetOutputFile = new File("/Users/vicwintriss/-League/Financial/Budget/Budget2021/Updated2021Budget.xlsx");
    public void writeBudget(XSSFWorkbook budgetWorkbook)
    {
        System.out.println("(9) Writing budget workbook to File: " + updated2020MasterBudgetOutputFile);
        try
        {
            FileOutputStream budgetOutputFOS = new FileOutputStream(updated2020MasterBudgetOutputFile);
            budgetWorkbook.write(budgetOutputFOS);
            budgetOutputFOS.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        System.out.println("(10) Finished writing budget workbook to File: " + updated2020MasterBudgetOutputFile);
    }
}