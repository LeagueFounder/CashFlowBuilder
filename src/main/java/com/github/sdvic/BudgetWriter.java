package com.github.sdvic;
/******************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190508
 * copyright 2019 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;

public class BudgetWriter
{
    public Workbook budgetWorkbook;
    public FileOutputStream budgetFOS;

    public BudgetWriter(Workbook budgetWorkbook, FileOutputStream budgetFOS)
    {
        this.budgetWorkbook = budgetWorkbook;
        this.budgetFOS = budgetFOS;
    }
    public void writeBudget()
    {
        try
        {
            budgetWorkbook.write(budgetFOS);
            budgetFOS.close();
        }
        catch (Exception e)
        {
            System.out.println("League budget write problems " + e);
        }
    }
}
