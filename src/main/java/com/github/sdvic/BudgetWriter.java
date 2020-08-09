package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200809
 * * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;

public class BudgetWriter
{
    private File updated2020MasterBudgetOutputFile = new File("/Users/VicMini/Desktop/Updated2020MasterBudgetOutputFile.xlsx");
    private FileOutputStream budgetOutputFOS;

    public void writeBudget(XSSFWorkbook budgetWorkbook)
    {
        try
        {
            budgetOutputFOS = new FileOutputStream(updated2020MasterBudgetOutputFile);
            budgetWorkbook.write(budgetOutputFOS);
            budgetOutputFOS.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        System.out.println("Finished writing budget workbook to File: " + updated2020MasterBudgetOutputFile);
    }
}
