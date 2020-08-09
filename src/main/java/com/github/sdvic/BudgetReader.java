package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200809
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

public class BudgetReader
{
    String budgetInputFileName = "/Users/VicMini/Desktop/Updated2020MasterBudgetOutputFile.xlsx";
    private File budgetInputFile;
    private FormulaEvaluator evaluator;
    private FileInputStream budgetInputFIS;
    private XSSFWorkbook budgetWorkBook;
    private XSSFSheet budgetSheet;
    private HashMap<String, Integer> budgetHashMap = new HashMap<>();
    private int budgetValue;
    private String budgetKey;
    private XSSFCell cell;

    public void readBudget()
    {
        try
        {
            budgetInputFile = new File(budgetInputFileName);
            budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            budgetInputFIS.close();
            budgetSheet = budgetWorkBook.getSheetAt(0);
        }
        catch (FileNotFoundException e)
        {
            System.out.println("file not found");
            e.printStackTrace();
        }
        catch (IOException e)
        {
            System.out.println("file IOexception");
            e.printStackTrace();
        }
        evaluator = budgetWorkBook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        for (Row row : budgetSheet)
        {
            for (int i = 0; i < 2; i++)
            {
                if (row.getCell(i) != null)
                {
                    switch (row.getCell(i).getCellType())
                    {
                        case XSSFCell.CELL_TYPE_BLANK://Type 3
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            break;
                        case XSSFCell.CELL_TYPE_FORMULA://Type 2
                            budgetValue = (int) row.getCell(i).getNumericCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            budgetValue = (int) row.getCell(i).getNumericCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_STRING://Type 1
                            budgetKey = row.getCell(i).getStringCellValue();
                            break;
                        default:
                            System.out.println("switch error");
                    }
                    budgetHashMap.put(budgetKey, budgetValue);
                }
            }
        }
        //budgetHashMap.forEach((K, V) -> System.out.println( K + " => " + V ));
        System.out.println("Finished reading Budget In budgetReader from " + "/Users/VicMini/Desktop/Updated2020MasterBudgetOutputFile.xlsx" + " to: budgetHashMap, HashMap size => " + budgetHashMap.size());
    }

    public HashMap<String, Integer> getBudgetHashMap()
    {
        return budgetHashMap;
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
