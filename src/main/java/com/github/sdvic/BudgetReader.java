package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200709
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class BudgetReader
{
    private final File budgetInputFile = new File("/Users/VicMini/Desktop/VicBudget2020.xlsx");
    private XSSFWorkbook budgetWorkBook;
    private XSSFCell row;
    private XSSFCell cell;
    private  FileInputStream budgetInputFIS;
    public void readBudget()
    {
        try
        {
            budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            budgetInputFIS.close();
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        //budgetSheet.forEach(row -> System.out.println( row.getCell(6) ));
    }

    public XSSFWorkbook getBudgetWorkBook()
    {
        return budgetWorkBook;
    }
}
