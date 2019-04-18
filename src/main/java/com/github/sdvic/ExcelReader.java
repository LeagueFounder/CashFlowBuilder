package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.4 April 17, 2019
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;

public class ExcelReader
{
    public HashMap<String, Integer> chartOfAccounts = new HashMap<>();
    public Sheet pandlSheet;
    public Sheet sarah5yearLocalSheet;

    public ExcelReader(Sheet pls, Sheet s5Ls)
    {
        System.out.println("Setting Up ==============");
        String cellKey = null;
        int cellValue = 0;
        String title = null;
        int amount = 0;
        pandlSheet = pls;
        sarah5yearLocalSheet = s5Ls;
        System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%% Cash Flow Analysis %%%%%%%%%%%%%%%%%%%%%%%");
        for (Row row : sarah5yearLocalSheet)//Bring full cash flow chart from Excel
        {
            if (row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
            {
                continue;
            }
            if (row.getCell(0).getCellType() == Cell.CELL_TYPE_STRING)
            {
                title = row.getCell(0).getStringCellValue();
                System.out.print(title + "         ");
            }
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC)
            {
                amount = (int) row.getCell(1).getNumericCellValue();
                System.out.println(amount + "\n");
            }
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_BLANK)
            {
                continue;
            }
            System.out.println();
        }
        System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Reading Chart Of Accounts &&&&&&&&&&&&&&&&&&&&&&&&&");
        for (Row row : pandlSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_BLANK)
            {
                continue;
            }
            if (row.getCell(0).getCellType() == Cell.CELL_TYPE_STRING)
            {
                cellKey = row.getCell(0).getStringCellValue();
                System.out.print(cellKey + "      ");
            }
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC)
            {
                cellValue = (int) row.getCell(1).getNumericCellValue();
                System.out.println(cellValue);
            }
            chartOfAccounts.put(cellKey, cellValue);
        }
    }
}
