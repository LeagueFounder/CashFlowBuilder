package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201119
// * copyright 2020 Vic Wintriss
//*******************************************************************************************
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
public class PandLReader
{
    private File pandLInputFile;
    private XSSFWorkbook pandLWorkBook;
    private XSSFCell valueCell;
    private XSSFCell keyCell;
    private String keyValue;
    private double valueValue;
    private final HashMap<String, Double> pandLMap = new HashMap<>();
    private String errMsg;
    private int cellType;
    /*******************************************************************************************************************
     * P&L Reader
     * Copies entire QuickBooks P&L to Hash Map
     ******************************************************************************************************************/
    public PandLReader(int targetMonth)
    {
        try
        {
            String inputFileName = "/Users/vicwintriss/-League/Financial/Budget/2020BudgetPandLs/" + targetMonth + "The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx";
            pandLInputFile = new File(inputFileName);
            System.out.println("(1) Started reading QuickBooks PandL Excel file from: " + pandLInputFile + " to pandlHashMap");
            FileInputStream pandlInputFIS = new FileInputStream(pandLInputFile);
            pandLWorkBook = new XSSFWorkbook(pandlInputFIS);
            pandlInputFIS.close();
        }
        catch (Exception e)
        {
            System.out.println("FileNotFoundException in readPandLtoHashMap()");
            e.printStackTrace();
        }
        FormulaEvaluator evaluator = pandLWorkBook.getCreationHelper().createFormulaEvaluator();
        XSSFFormulaEvaluator.evaluateAllFormulaCells(pandLWorkBook);
        XSSFSheet pandlSheet = pandLWorkBook.getSheetAt(0);
        for (int rowIndex = 0; rowIndex < pandlSheet.getLastRowNum(); rowIndex++)
        {
            XSSFRow row = pandlSheet.getRow(rowIndex);
            if (row == null)
            {
                continue;
            }
            XSSFCell keyCell = row.getCell(0); //Key cell
            if (keyCell == null)
            {
                continue;
            }
            try
            {
                keyValue = keyCell.getStringCellValue();
                final CellValue keyCellValue = evaluator.evaluate(keyCell);
                String keyStringRaw = ((CellValue) keyCellValue).formatAsString().trim();//Found Key String
                keyValue = keyStringRaw.replaceAll("^\"+|\"+$", "").trim();//Strip off quote signs
                keyValue.trim();
            }
            catch (Exception e)
            {
                System.out.println("Can't get key String value at line 65");
                continue;
            }
            XSSFCell valueCell = row.getCell(1); //Value cell
            if (valueCell == null)
            {
                continue;
            }
            try
            {
                valueValue = valueCell.getNumericCellValue();
            }
            catch(Exception e)
            {
                System.out.println("Can't get numeric value at line 80");
            }
            getPandLMap().put(keyValue, valueValue);
        }
//        System.out.println("        ===========PandL Map======================");
//        getPandLMap().forEach((K, V) -> System.out.println("             " +  K + " => " + V ));
        System.out.println("(2) Finished reading QuickBooks PandL Excel file from: " + pandLInputFile + " to: pandlHashMap, HashMap size: " + getPandLMap().size());
    }
    public HashMap<String, Double> getPandLMap()
    {
        return pandLMap;
    }
    public XSSFWorkbook getPandlWorkBook()
    {
        return pandLWorkBook;
    }
}