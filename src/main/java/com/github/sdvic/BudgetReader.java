package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201029
// * copyright 2020 Vic Wintriss
//*******************************************************************************************
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
public class BudgetReader
{
    String budgetInputFileName = "/Users/vicwintriss/-League/Financial/Budget/2020BudgetPandLs/SarahRevisedBudget2020.xlsx";
    String updateInputFileName = "/Users/vicwintriss/-League/Financial/Budget/2020BudgetPandLs/Updated2020MasterBudgetOutputFile.xlsx";
    private XSSFWorkbook budgetWorkBook;
    private final HashMap<String, Integer> budgetMap = new HashMap<>();
    private XSSFCell cell;
    private XSSFCell keyCell;
    private XSSFCell valueCell;
    private String followOnAnswer;
    public void readBudget(int targetMonth, String followOnAnswer)
    {
        System.out.println("(3) Starting reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap");
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
        for (int rowIndex = 0; rowIndex < budgetSheet.getLastRowNum(); rowIndex++)
        {
            XSSFRow row = budgetSheet.getRow(rowIndex);
            Cell keyCell = row.getCell(0);
            String keyString = "";
            int switcher = keyCell.getCellType();
            switch (switcher)//Read Budget Key From Excel Spreadsheet
            {
                case 0://NUMERIC
//                    System.out.println("***Key error 63 while reading Budget Excel spreadsheet key...found XSSF cell type number instead of String =>  \"" + row.getCell(0) + "\"" + " at row " + (rowIndex + 1));
                    break;
                case 1://STRING
                    final CellValue keyCellValue = evaluator.evaluate(keyCell);
                    keyString = keyCellValue.formatAsString().trim();//Found Key String
                    break;
                case 2://FORMULA
//                    System.out.println("***Key error 69 while reading Budget Excel spreadsheet key...found XSSF cell type formula instead of String =>  \"" + row.getCell(0) + "\"" + " at row " + (rowIndex + 1));
                    break;
                case 3://BOOLEAN
//                    System.out.println("***Key Error 72 while reading Budget Excel spreadsheet key...found XSSF cell type boolean instead of String =>  \"" + row.getCell(0) + "\"" + " at row " + (rowIndex + 1));
                case 4://ERROR
//                    System.out.println("***Key error 75 while reading Budget Excel spreadsheet key...found XSSF cell type ERROR instead of String =>  \"" + row.getCell(0) + "\"" + " at row " + (rowIndex + 1));
                    break;
                default:
                    break;
            }
            Cell valueCell = row.getCell(targetMonth);
            int valueInt = -1;
            switcher = valueCell.getCellType();
            switch (switcher)//Read Budget Value From Excel Spreadsheet
            {
                case 0://NUMERIC
                    valueInt = (int) row.getCell(targetMonth).getNumericCellValue();
                    break;
                case 1://STRING
//                    System.out.println("***Value error 90 while reading Budget Excel spreadsheet...found XSSF cell type String => \"" + row.getCell(targetMonth)  + "\" instead of number at row " + (rowIndex + 1));
                    break;

                case 2://FORMULA
//                    System.out.println("***Value error 93 while reading Budget Excel spreadsheet value...found XSSF cell type formula instead of number =>  \"" + row.getCell(targetMonth) + "\"" + " at row " + (rowIndex + 1));
                    break;
                case 3://BOOLEAN
//                    System.out.println("***Value error 96 while reading Budget Excel spreadsheet value...found XSSF cell type boolean instead of number =>  \"" + row.getCell(targetMonth) + "\"" + " at row " + (rowIndex + 1));
                    break;

                case 4://ERROR
//                    System.out.println("***Value error 99 while reading Budget Excel spreadsheet value...found XSSF cell type error instead of number =>  \"" + row.getCell(targetMonth) + "\"" + " at row " + (rowIndex + 1));
                    break;
                default:
                    break;
            }
            budgetMap.put(keyString, valueInt);
        }
//        System.out.println("           Budget Map For Month " + targetMonth);
//        budgetMap.forEach((K, V) -> System.out.println("      " + K + " => " + V));
        System.out.println("(4) Finished reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap, HashMap size: " + budgetMap.size());
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