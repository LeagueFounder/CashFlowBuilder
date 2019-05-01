package com.github.sdvic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;

public class CashFlowItemAggregator
{
    public CashFlowItemAggregator(Workbook sarah5yearWorkbook, HashMap<String, Integer> chartOfAccountsMap, String version)
    {
        Cell revDateCell = sarah5yearWorkbook.getSheetAt(0).getRow(1).getCell(0);
        revDateCell.setCellValue(version);
        Cell cashContributions = sarah5yearWorkbook.getSheetAt(0).getRow(1).getCell(6);
        System.out.println("Chart Of Acclounts HashMap size => " + chartOfAccountsMap.size());
        System.out.println("43450 => " + chartOfAccountsMap.get("      43450 Individ, Business Contributions"));
    }
}
