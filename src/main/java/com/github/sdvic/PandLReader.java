package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200820
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
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
    private HashMap<String, Integer> pandlHashMap = new HashMap<>();
    private FormulaEvaluator evaluator;

    /*******************************************************************************************************************
     * P&L Reader
     * Copies entire QuickBooks P&L to Hash Map
     ******************************************************************************************************************/
    public void readPandL(int targetMonth)
    {
        System.out.println("(1) Started reading PandL In PandLreader from: " + pandInputlFile + " to: pandlHashMap, HashMap size: " + pandlHashMap.size());
        try
        {
            inputFileName = "/Users/VicMini/Desktop/" + targetMonth + "The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx";
            pandInputlFile = new File(inputFileName);
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
        for (Row row : pandlSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            for (Cell cell : row)
            {
                switch (cell.getCellType())
                    {
                        case XSSFCell.CELL_TYPE_BLANK://Type 3
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            break;
                        case XSSFCell.CELL_TYPE_FORMULA://Type 2
                            cellValue = (int)cell.getNumericCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            cellValue = (int) cell.getNumericCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_STRING://Type 1
                            cellKey = cell.getStringCellValue().trim();
                            break;
                        default:
                            System.out.println("switch error");
                    }
                pandlHashMap.put(cellKey, cellValue);
            }
        }
        //pandlHashMap.forEach((K, V) -> System.out.println( K + " => " + V ));
        System.out.println("(2) Finished reading PandL In PandLreader from: " + pandInputlFile + " to: pandlHashMap, HashMap size: " + pandlHashMap.size());
    }
    public HashMap<String, Integer> getPandlHashMap()
    {
        return pandlHashMap;
    }
    public XSSFWorkbook getPandlWorkBook()
    {
        return pandlWorkBook;
    }
}
