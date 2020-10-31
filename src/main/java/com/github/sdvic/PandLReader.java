package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201031
// * copyright 2020 Vic Wintriss
//*******************************************************************************************
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
    private int valueValue;
    private final HashMap<String, Integer> pandLMap = new HashMap<>();
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
            XSSFCell valueCell = row.getCell(1); //Value cell
            if (valueCell == null)
            {
                continue;
            }
            if (keyCell.getCellType() == XSSFCell.CELL_TYPE_STRING)
            {
                keyValue = keyCell.getStringCellValue().trim();//Found Key String
            }
            else
            {
                keyValue = "No Key found";
                continue;
            }
            if (valueCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
            {
                valueValue = (int) valueCell.getNumericCellValue();
            }
            else
            {
                valueValue = -1;
            }

            getPandLMap().put(keyValue, valueValue);
        }
//        System.out.println("        ===========PandL Map======================");
//        pandlMap.forEach((K, V) -> System.out.println("             " +  K + " => " + V ));
        System.out.println("(2) Finished reading QuickBooks PandL Excel file from: " + pandLInputFile + " to: pandlHashMap, HashMap size: " + getPandLMap().size());
    }
    public HashMap<String, Integer> getPandLMap()
    {
        return pandLMap;
    }
    public XSSFWorkbook getPandlWorkBook()
    {
        return pandLWorkBook;
    }
}