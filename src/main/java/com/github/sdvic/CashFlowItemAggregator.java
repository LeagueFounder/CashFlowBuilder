package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190502
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import java.util.HashMap;

public class CashFlowItemAggregator
{
    public int cashContributions;
    int cashTuition;
    int totalCashIncome;
    int payrollExpense;

    public CashFlowItemAggregator(Workbook sarah5yearWorkbook, HashMap<String, Integer> pandLmap, String version)
    {
        Sheet fiveYearSheet = sarah5yearWorkbook.getSheetAt(0);
        Cell revDateCell = fiveYearSheet.getRow(1).getCell(0);
        revDateCell.setCellValue(version);
        for (Row row : fiveYearSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            Cell cell = row.getCell(0);
            if (cell.getCellType() != Cell.CELL_TYPE_STRING)
            {
                continue;
            }
            else
            {
                switch (cell.getStringCellValue())
                {
                    case "Cash Contributions":
                        Cell cashContributionsCell = fiveYearSheet.getRow(3).getCell(6);
                        cashContributions = pandLmap.get("      43450 Individ, Business Contributions");
                        cashContributionsCell.setCellValue(cashContributions);
                        System.out.println(cell.getStringCellValue() + " => " + cashContributions);
                        break;
                    case "Cash Tuition ":
                        Cell cashTuitionCell = fiveYearSheet.getRow(5).getCell(6);
                        cashTuition = pandLmap.get("      Total 47201 Tuition  Fees");
                        cashTuitionCell.setCellValue(cashTuition);
                        System.out.println(cell.getStringCellValue() + " => " + cashTuition);
                        break;
                    case "Total Cash Income":
                        Cell totalCashIncomeCell = fiveYearSheet.getRow(6).getCell(6);
                        totalCashIncome = cashContributions + cashTuition;
                        totalCashIncomeCell.setCellValue(totalCashIncome);
                        System.out.println(cell.getStringCellValue() + " => " + totalCashIncome);
                        break;
                    case "Payroll":
                        Cell payrollCell = fiveYearSheet.getRow(9).getCell(6);
                        payrollExpense = pandLmap.get("   Total 62000 Salaries & Related Expenses");
                        payrollCell.setCellValue(payrollExpense);
                        System.out.println(cell.getStringCellValue() + " => " + payrollExpense);
                        break;
                    default:
                }
            }
        }
    }
}



