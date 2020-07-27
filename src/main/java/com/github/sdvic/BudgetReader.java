package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200726
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

public class BudgetReader
{
    private final File budgetInputFile = new File("/Users/VicMini/Desktop/PurpleBudget2020.xlsx");
    private FileInputStream budgetInputFIS;
    private CellStyle backgroundStyle;
    private XSSFWorkbook budgetWorkbook;
    private XSSFWorkbook budgetWorkBook;
    private XSSFRow row;
    private XSSFSheet currentBudgetSheet;
    private XSSFSheet currentSheet;
    public void readBudget()
    {
        try
        {
            budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            currentSheet = budgetWorkBook.getSheetAt(0);
            budgetInputFIS.close();
            currentBudgetSheet = budgetWorkBook.getSheetAt(0);
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
        System.out.println();
        System.out.println("Finished reading budget in BudgetReader from File " + budgetInputFile + " sheet size => " + budgetWorkBook.getSheetAt(0).getLastRowNum());
    }

    public XSSFWorkbook getBudgetWorkBook()
    {
        return budgetWorkBook;
    }
}
