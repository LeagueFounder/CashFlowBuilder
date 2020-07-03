package com.github.sdvic;
/******************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190508
 * copyright 2019 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.HashMap;

public class ExcelReader
{
    private XSSFWorkbook budgetWorkBook;
    private XSSFWorkbook pandlWorkBook;
    private XSSFSheet pandlSheet;
    private XSSFSheet budgetSheet;
    private String version;
    private String cellKey = null;
    private int cellValue = 0;
    private HashMap<String, Integer> pandlMap = new HashMap<>();
    private CellValue cv;
    private FormulaEvaluator evaluator;

    ///*******************************************************************************************************************
    //     * Excel Budget Writer
    //     * Writes appropriate, aggregated P & L items from Hash Map to Excel Budget Workbook
    //     *******************************************************************************************************************/
    public ExcelReader(XSSFWorkbook pandlWorkBook, XSSFWorkbook budgetWorkBook, String version)
    {
        this.budgetWorkBook = budgetWorkBook;
        this.pandlWorkBook = pandlWorkBook;
        this.version = version;
        readPandLtoHashMap(pandlWorkBook);
        readBudget(budgetWorkBook);
    }

    private void readBudget(XSSFWorkbook budgetWorkBook)
    {
        budgetSheet = budgetWorkBook.getSheetAt(0);
        evaluator = budgetWorkBook.getCreationHelper().createFormulaEvaluator();
    }
    /*******************************************************************************************************************
     * P&L Reader
     * Copies entire QuickBooks P&L to Hash Map
     *******************************************************************************************************************/
    private HashMap<String, Integer> readPandLtoHashMap(XSSFWorkbook pandlWorkBook)
    {
        pandlMap.clear();
        pandlSheet = pandlWorkBook.getSheetAt(0);
        evaluator = pandlWorkBook.getCreationHelper().createFormulaEvaluator();
        for (Row row : pandlSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            for (Cell cell : row)
            {
                cv = evaluator.evaluate(cell);
                    switch (cell.getCellType())
                    {
                        case XSSFCell.CELL_TYPE_BLANK:
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            break;
                        case XSSFCell.CELL_TYPE_FORMULA://Type 2
                            cellValue = (int)cv.getNumberValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            cellValue = (int) cell.getNumericCellValue();
                            System.out.println("*Unexpectedly getting numeric cell from QuickBooks Excel P&L file");
                            break;
                        case XSSFCell.CELL_TYPE_STRING://Type 1
                            cellKey = cv.getStringValue();
                            break;
                        default:
                            System.out.println("switch error");
                    }
                pandlMap.put(cellKey, cellValue);
            }
        }
//        tempMap.forEach((K, V) -> System.out.println( K + " => " + V ));
        return pandlMap;
    }
    public HashMap<String, Integer> getPandLMap()
    {
        return pandlMap;
    }
}
