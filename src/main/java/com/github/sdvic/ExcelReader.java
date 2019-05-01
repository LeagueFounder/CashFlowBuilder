package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.6 April 27, 2019
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

public class ExcelReader
{
    private HashMap<String, Integer> chartOfAccountsMap = new HashMap<>();
    public Sheet pandlSheet;
    public Workbook sarah5yearLocalWorkbook = new XSSFWorkbook();
    public XSSFSheet sarah5yearLocalSheet;
    public FormulaEvaluator plEvaluator;
    public String version;

    public ExcelReader(Workbook pandlWorkBook, Workbook s5yrwb, String version)
    {
        this.version = version;
        String cellKey = null;
        int cellValue = 0;
        pandlSheet = pandlWorkBook.getSheetAt(0);
        plEvaluator = pandlWorkBook.getCreationHelper().createFormulaEvaluator();
        sarah5yearLocalWorkbook = (XSSFWorkbook) s5yrwb;
        sarah5yearLocalSheet = (XSSFSheet) sarah5yearLocalWorkbook.getSheetAt(0);
        System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Reading Chart Of Accounts &&&&&&&&&&&&&&&&&&&&&&&&&");
        for (Row row : pandlSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            {
                for (Cell cell : row)
                {
                    {
                        plEvaluator.evaluateAll();
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
        }

//        for (Map.Entry<String, Integer> pair : chartOfAccountsMap.entrySet())
//        {
//            System.out.println(pair.getKey() + "   " + (int)pair.getValue());
//        }
        System.out.println("========================================end chart of accounts======================");
        System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%% Cash Flow Analysis %%%%%%%%%%%%%%%%%%%%%%%");
        plEvaluator.evaluateAll();
        //System.out.println(new XSSFExcelExtractor((XSSFWorkbook) sarah5yearLocalWorkbook).getText());
    }

    public HashMap<String, Integer> getChartOfAccountsMap()
    {
        return chartOfAccountsMap;
    }
}
