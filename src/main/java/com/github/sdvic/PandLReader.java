package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 200905
// * copyright 2020 Vic Wintriss
//******************************************************************************************

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;

public class PandLReader
{
    private String inputFileName;
    private File pandInputlFile;
    private FileInputStream pandlInputFIS;
    private XSSFWorkbook pandlWorkBook;
    private XSSFSheet pandlSheet;
    private String cellKey;
    private int cellValue;
    private HashMap<String, Integer> pandlMap = new HashMap<>();
    private FormulaEvaluator evaluator;

    /*******************************************************************************************************************
     * P&L Reader
     * Copies entire QuickBooks P&L to Hash Map
     ******************************************************************************************************************/
    public void readPandL(int targetMonth)
    {
        try
        {
            inputFileName = "/Users/VicMini/Desktop/" + targetMonth + "The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx";
            pandInputlFile = new File(inputFileName);
            System.out.println("(1) Started reading QuickBooks PandL Excel file from: " + pandInputlFile + " to pandlHashMap");
            pandlInputFIS = new FileInputStream(pandInputlFile);
            pandlWorkBook = new XSSFWorkbook(pandlInputFIS);
            pandlInputFIS.close();
        }
        catch (Exception e)
        {
            System.out.println("FileNotFoundException in readPandLtoHashMap()");
            e.printStackTrace();
        }
        evaluator = pandlWorkBook.getCreationHelper().createFormulaEvaluator();
        XSSFFormulaEvaluator.evaluateAllFormulaCells(pandlWorkBook);
        pandlSheet = pandlWorkBook.getSheetAt(0);
        for (int row = 0; row < pandlSheet.getLastRowNum(); row++)
        {
            if (pandlSheet.getRow(row) != null)
            {
                if (pandlSheet.getRow(row).getCell(0) != null)
                {
                    Cell keyCell = pandlSheet.getRow(row).getCell(0);
                    Cell valueCell = pandlSheet.getRow(row).getCell(1);
                    if (keyCell.getCellType() == XSSFCell.CELL_TYPE_STRING)
                    {
                        if (valueCell.getCellType() == XSSFCell.CELL_TYPE_FORMULA)
                        {
                            cellKey = keyCell.getStringCellValue().trim();
                            cellValue = (int) valueCell.getNumericCellValue();
                            pandlMap.put(cellKey, cellValue);
                        }
                    }
                }
            }
        }
        pandlMap.forEach((K, V) -> System.out.println( K + " => " + V ));
        System.out.println("(2) Finished reading QuickBooks PandL Excel file from: " + pandInputlFile + " to: pandlHashMap, HashMap size: " + pandlMap.size());
    }

    public HashMap<String, Integer> getPandlMap()
    {
        return pandlMap;
    }

    public XSSFWorkbook getPandlWorkBook()
    {
        return pandlWorkBook;
    }
}