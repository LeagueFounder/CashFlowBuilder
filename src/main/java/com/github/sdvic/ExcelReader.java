package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.6 April 27, 2019
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.HashMap;
import java.util.Map;

public class ExcelReader
{
    public HashMap<String, Integer> chartOfAccountsMap = new HashMap<>();
    public Sheet pandlSheet;
    public Sheet sarah5yearSheet;
    public FormulaEvaluator evaluator;

    public ExcelReader(Sheet pls, Sheet s5Ls, Workbook s5yrwb)
    {
        System.out.println("Setting Up ==============");
        String cellKey = null;
        int cellValue = 0;
        pandlSheet = pls;
        sarah5yearSheet = s5Ls;
        evaluator = s5yrwb.getCreationHelper().createFormulaEvaluator();

        System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Reading Chart Of Accounts &&&&&&&&&&&&&&&&&&&&&&&&&");
        for (Row row : pls)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            if (isRowEmpty(row))
            {
                continue;
            }
            else
            {
                for (Cell cell : row)
                {
                    if (cell != null)
                    {
                        evaluator.evaluateAll();
                        switch (cell.getCellType())
                        {
                            case XSSFCell.CELL_TYPE_BLANK:
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                break;
                            case XSSFCell.CELL_TYPE_FORMULA:
                                System.out.print((int)cell.getNumericCellValue() + "   ");
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print((int)cell.getNumericCellValue() + "\t\t\t\t");
                                break;
                            case XSSFCell.CELL_TYPE_STRING:
                                System.out.print(cell + "\t");
                                break;
                        }
                        System.out.print("\t");
                    }
                }
                System.out.println();
            }
            chartOfAccountsMap.put(cellKey, cellValue);
        }
        System.out.println("Chart Of Acclounts HashMap size => " + chartOfAccountsMap.size());
        for (Map.Entry<String, Integer> pair : chartOfAccountsMap.entrySet())
        {
            System.out.println(pair.getKey() + "   " + pair.getValue());

        }
        System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%% Cash Flow Analysis %%%%%%%%%%%%%%%%%%%%%%%");
        evaluator.evaluateAll();
        for (Row row : sarah5yearSheet)//Bring full cash flow chart from Excel
        {
            if (isRowEmpty(row))
            {
                continue;
            }
            else
            {
                for (Cell cell : row)
                {
                    if (cell != null)
                    {
                        switch (cell.getCellType())
                        {
                            case XSSFCell.CELL_TYPE_BLANK:
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                break;
                            case XSSFCell.CELL_TYPE_FORMULA:
                                System.out.print((int)cell.getNumericCellValue() + "   ");
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print((int)cell.getNumericCellValue() + "   ");
                                break;
                            case XSSFCell.CELL_TYPE_STRING:
                                System.out.print(cell + "\t");
                                break;
                        }
                        System.out.print("\t");
                    }
                }
                System.out.println();
            }
        }
    }

    public static boolean isRowEmpty(Row row)//From StackOverflow
    {
        boolean isEmpty = true;
        DataFormatter dataFormatter = new DataFormatter();
        if (row != null)
        {

            for (Cell cell : row)
            {
                if (dataFormatter.formatCellValue(cell).trim().length() > 0)
                {
                    isEmpty = false;
                    break;
                }
            }
        }
        return isEmpty;
    }
}
