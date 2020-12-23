package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201222
// * copyright 2020 Vic Wintriss
//*******************************************************************************************
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.concurrent.atomic.AtomicInteger;
public class BudgetReader
{
    String budgetInputFileName = "/Users/vicwintriss/-League/Financial/Budget/2020BudgetPandLs/SarahRevisedBudget2020.xlsx";
    String updateInputFileName = "/Users/vicwintriss/-League/Financial/Budget/2020BudgetPandLs/Updated2020MasterBudgetOutputFile.xlsx";
    private XSSFWorkbook budgetWorkBook;
    private final HashMap<String, Integer> budgetMap = new HashMap<>();
    public BudgetReader(int targetMonth, String followOnAnswer)
    {
        System.out.println("(3) Starting reading Budget In budgetReader Constructor from " + budgetInputFileName + " to: budgetHashMap");
        XSSFSheet budgetSheet;
        try
        {
            File budgetInputFile;
            if (followOnAnswer.equals("Yes"))
            {
                budgetInputFile = new File(budgetInputFileName);
            }
            else
            {
                budgetInputFile = new File(updateInputFileName);
            }
            FileInputStream budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            budgetInputFIS.close();
        }
        catch (Exception e)
        {
            System.out.println("Exception 48 reading budget");
            e.printStackTrace();
        }
        FormulaEvaluator evaluator = budgetWorkBook.getCreationHelper().createFormulaEvaluator();
        XSSFFormulaEvaluator.evaluateAllFormulaCells(budgetWorkBook);
        budgetSheet = budgetWorkBook.getSheetAt(0);
        System.out.println("budget sheet => " + budgetSheet);
        System.out.println("Last budget row number is => " + budgetSheet.getLastRowNum());
        for (int rowIndex = 0; rowIndex < 34; rowIndex++)
        {
            XSSFRow budgetRow = budgetSheet.getRow(rowIndex);
            Cell keyCell = budgetRow.getCell(0);
            String keyString = "";
            keyString = keyCell.getStringCellValue();
            keyString = keyString.replaceAll("^\"+|\"+$", "");//Strip off quote signs
            Cell valueCell = budgetRow.getCell(targetMonth);
            int valueInt = -1;
            try
            {
                valueInt = (int) valueCell.getNumericCellValue();
            }
            catch (Exception e)
            {

            }
            getBudgetMap().put(keyString, valueInt);
        }
//        System.out.println("           Budget Map For Month " + targetMonth);
//        budgetMap.forEach((K, V) -> System.out.println("      " + K + " => " + V));
        System.out.println("(4) Finished reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap, HashMap size: " + getBudgetMap().size());
    }
    public HashMap<String, Integer> getBudgetMap()
    {
        return budgetMap;
    }
    public XSSFWorkbook getBudgetWorkBook()
    {
        return budgetWorkBook;
    }
    public void setBudgetWorkBook(XSSFWorkbook budgetWorkBook)
    {
        this.budgetWorkBook = budgetWorkBook;
    }
}