package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200709
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class BudgetWriter
{
    private File budgetOutputFile = new File("/Users/VicMini/Desktop/VicBudget2020Mod.xlsx");
    private FileOutputStream budgetOutputFOS;
    private XSSFWorkbook budgetWorkbook;
    public void writeBudget(XSSFWorkbook budgetWorkbook)
    {
        this.budgetWorkbook = budgetWorkbook;
        try
        {
            budgetOutputFOS = new FileOutputStream(budgetOutputFile);
            budgetWorkbook.write(budgetOutputFOS);
            budgetOutputFOS.close();
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }
}
