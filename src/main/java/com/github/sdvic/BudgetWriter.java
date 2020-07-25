package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200725
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
public class BudgetWriter
{
    private File updated2020MasterBudgetOutputFile = new File("/Users/VicMini/Desktop/Updated2020MasterBudgetOutputFile.xlsx");
    private FileOutputStream budgetOutputFOS;
    private XSSFWorkbook budgetWorkbook;
    public void writeBudget(XSSFWorkbook budgetWorkbook)
    {
        this.budgetWorkbook = budgetWorkbook;
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
