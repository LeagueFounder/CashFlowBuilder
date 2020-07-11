package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200709
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;

public class PandLReader
{
    private File pandInputlFile = new File("/Users/VicMini/Desktop/The+League+of+Amazing+Programmers_Profit+and+Loss.xlsx");
    private FileInputStream pandlInputFIS;
    private XSSFWorkbook pandlWorkBook;
    private XSSFSheet pandlSheet;
    private String cellKey = null;
    private int cellValue = 0;
    private HashMap<String, Integer> pandlHashMap = new HashMap<>();
    /*******************************************************************************************************************
     * P&L Reader
     * Copies entire QuickBooks P&L to Hash Map
     ******************************************************************************************************************/
    public void readPandLtoHashMap()
    {
        try
        {
            pandlInputFIS = new FileInputStream(pandInputlFile);
            pandlWorkBook = new XSSFWorkbook(pandlInputFIS);
            pandlInputFIS.close();
        }
        catch (FileNotFoundException e)
        {
            System.out.println("FileNotFoundException in readPandLtoHashMap()");
            e.printStackTrace();
        }
        catch (IOException e)
        {
            System.out.println("IOException in readPandLtoHashMap()");
            e.printStackTrace();
        }
        pandlSheet = pandlWorkBook.getSheetAt(0);
        for (Row row : pandlSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            for (Cell cell : row)
            {
                switch (cell.getCellType())
                    {
                        case XSSFCell.CELL_TYPE_BLANK:
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
    }

    public HashMap<String, Integer> getPandlHashMap()
    {
        return pandlHashMap;
    }
}
