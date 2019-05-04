package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190504
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

public class ExcelReader
{
    public Sheet pandlSheet;
    public Workbook sarah5yearLocalWorkbook = new XSSFWorkbook();
    public XSSFSheet sarah5yearLocalSheet;
    public String version;
    private HashMap<String, Integer> chartOfAccountsMap = new HashMap<>();

    public ExcelReader(Workbook pandlWorkBook, Workbook s5yrwb, String version)
    {
        this.version = version;
        String cellKey = null;
        int cellValue = 0;
        pandlSheet = pandlWorkBook.getSheetAt(0);
        sarah5yearLocalWorkbook = (XSSFWorkbook) s5yrwb;
        sarah5yearLocalSheet = (XSSFSheet) sarah5yearLocalWorkbook.getSheetAt(0);
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
                    case XSSFCell.CELL_TYPE_FORMULA:
                        cellValue = (int) cell.getNumericCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        cellValue = (int) cell.getNumericCellValue();
                        break;
                    case XSSFCell.CELL_TYPE_STRING:
                        cellKey = cell.getStringCellValue();
                        break;
                    default:
                        System.out.println("switch error");
                }
                chartOfAccountsMap.put(cellKey, cellValue);
            }
        }
    }

    public HashMap<String, Integer> getChartOfAccountsMap()
    {
        return chartOfAccountsMap;
    }
}
