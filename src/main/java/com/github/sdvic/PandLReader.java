package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 200918
// * copyright 2020 Vic Wintriss
//******************************************************************************************

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.*;

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
    private XSSFRow row;
    private XSSFCell valueCell;
    private XSSFCell keyCell;
    private String keyValue;
    private int valueValue;
    private HashMap<String, Integer> pandlMap = new HashMap<>();
    private FormulaEvaluator evaluator;
    private String errMsg;
    private int cellType;
    private String switcher;

    /*******************************************************************************************************************
     * P&L Reader
     * Copies entire QuickBooks P&L to Hash Map
     ******************************************************************************************************************/
    public void readPandL(int targetMonth)
    {
        try
        {
            inputFileName = "/Users/vicwintriss/Desktop/" + targetMonth + "The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx";
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
        for (int rowIndex = 0; rowIndex < pandlSheet.getLastRowNum(); rowIndex++)
        {
            row = pandlSheet.getRow(rowIndex);
           try
           {
               switcher = Integer.toString(row.getCell(0).getCellType());
           }catch (Exception e)
            {
               switcher = "null";
            }
        switch (switcher)//PandL Key
            {
                case "0"://NUMERIC
                    System.out.println("Error reading PandL sheet...found number...looking for PandL key string");
                    break;
                case "1"://STRING
                    keyValue = row.getCell(0).getStringCellValue().trim();//Key
                    break;
                case "2"://FORMULA
                    System.out.println("Error reading PandL sheet...found formula...looking for PandL key");
                    break;
                case "3":// BLANK
                    System.out.println("Error reading PandL sheet...found blank PandL key");
                    break;
                case "4"://BOOLEAN
                    System.out.println("Error reading PandL sheet...found boolean PandL key");
                    break;
                case "5"://ERROR
                    System.out.println("Error reading PandL sheet...found XSSF cell type ERROR PandL key");
                    break;
                case "null":
                    System.out.println("Error reading PandL sheet...found null PandL key");
                    break;
                default:
                    System.out.println("Error reading PandL sheet...found switch default PandL key");
                    break;
            }
            try
            {
                switcher = Integer.toString(row.getCell(1).getCellType());
            }catch (Exception e)
            {
                switcher = "null";
            }
            switch (switcher)//PanL Value
            {
                case "0"://NUMERIC
                    valueValue = (int) row.getCell(1).getNumericCellValue();//Value
                    break;
                case "1"://STRING
                    System.out.println("Error reading PandL sheet...found string PandL value");
                    break;
                case "2"://FORMULA
                    valueValue = (int) row.getCell(1).getNumericCellValue();//value
                    break;
                case "3":// BLANK
                    System.out.println("Error reading PandL sheet...found blank PandL value");
                    break;
                case "4"://BOOLEAN
                    System.out.println("Error reading PandL sheet...found boolean PandL value");
                    break;
                case "5"://ERROR
                    System.out.println("Error reading PandL sheet...found XSSF cell type ERROR PandL value");
                    break;
                case "null":
                    System.out.println("Error reading PandL sheet...found null PandL value");
                    break;
                default:
                    System.out.println("switch default while reading PandL sheet value");
                    break;
            }
            pandlMap.put(keyValue, valueValue);
        }
        //pandlMap.forEach((K, V) -> System.out.println( K + " => " + V ));
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