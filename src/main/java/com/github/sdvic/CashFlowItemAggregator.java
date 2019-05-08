package com.github.sdvic;
/******************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190508
 * copyright 2019 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.*;
import java.util.HashMap;

public class CashFlowItemAggregator
{
    public CashFlowItemAggregator(Workbook sarah5yearWorkbook, HashMap<String, Integer> pandLmap, String version, int yearColumnIndex)
    {
        Sheet fiveYearSheet = sarah5yearWorkbook.getSheetAt(0);
        Cell revDateCell = fiveYearSheet.getRow(1).getCell(0);
        revDateCell.setCellValue(version);
        for (Row row : fiveYearSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            Cell fiveYearProjectionLabelCell = row.getCell(0);
            if (fiveYearProjectionLabelCell.getCellType() == Cell.CELL_TYPE_STRING)
            {
                switch (fiveYearProjectionLabelCell.getStringCellValue())
                {
                    case "Cash Grants":
                        Cell cashGrantsCell = fiveYearSheet.getRow(4).getCell(yearColumnIndex);
                        cashGrantsCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        int cashGrants = 0;
                        if (pandLmap.get("      Total 47204 Grant Scholarship") != null)
                        {
                            cashGrants = (pandLmap.get("      Total 47204 Grant Scholarship"));
                        }
                        cashGrantsCell.setCellValue(cashGrants);
                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + cashGrants);
                        break;
                    case "Cash Contributions":
                        Cell cashContributionsCell = fiveYearSheet.getRow(3).getCell(yearColumnIndex);
                        cashContributionsCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        int cashContributions = 0;
                        if (pandLmap.get("      43450 Individ, Business Contributions") != null)
                        {
                            cashContributions = pandLmap.get("      43450 Individ, Business Contributions");
                        }
                        cashContributionsCell.setCellValue(cashContributions);
                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + cashContributions);
                        break;
                    case "Cash Tuition ":
                        Cell cashTuitionCell = fiveYearSheet.getRow(5).getCell(yearColumnIndex);
                        cashTuitionCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        int cashTuition = 0;
                        int workshopFees = 0;
                        int totalTuition = 0;
                        if (pandLmap.get("      Total 47201 Tuition  Fees") != null)
                        {
                            cashTuition = pandLmap.get("      Total 47201 Tuition  Fees");
                        }
                        if (pandLmap.get("      Total 47202 Workshop Fees") != null)
                        {
                            workshopFees = pandLmap.get("      Total 47202 Workshop Fees");
                        }
                        totalTuition = cashTuition + workshopFees;
                        cashTuitionCell.setCellValue((int) (totalTuition));
                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + totalTuition);
                        break;
                    case "Payroll":
                        Cell payrollCell = fiveYearSheet.getRow(9).getCell(yearColumnIndex);
                        payrollCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        int payrollExpense = pandLmap.get("   Total 62000 Salaries & Related Expenses");
                        int payrollFees = pandLmap.get("   62145 Payroll Service Fees");
                        int contributedServices = pandLmap.get("      62010 Salaries contributed services");
                        int payroll = payrollExpense + payrollFees - contributedServices;
                        payrollCell.setCellValue(payroll);
                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + payroll);
                        break;
                    case "Facilities":
                        Cell facilitiesCell = fiveYearSheet.getRow(10).getCell(yearColumnIndex);
                        facilitiesCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        int facilityExpense = pandLmap.get("   Total 62800 Facilities and Equipment");
                        facilitiesCell.setCellValue(facilityExpense);
                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + facilityExpense);
                        break;
                    case "All Other Expenses":
                        Cell otherExpenseCell = fiveYearSheet.getRow(11).getCell(yearColumnIndex);
                        otherExpenseCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        int operationsExpense = pandLmap.get("   Total 65000 Operations");
                        int otherExpense = pandLmap.get("   Total 65100 Other Types of Expenses");
                        int travelExpense = pandLmap.get("   Total 68300 Travel and Meetings");
                        int contractServiceExpense = 0;
                        int allOtherExpenses = 0;
                        if (pandLmap.get("   Total 62100 Contract Services") != null)
                        {
                            otherExpenseCell.setCellValue(allOtherExpenses);
                        }
                        allOtherExpenses = operationsExpense + otherExpense + travelExpense + contractServiceExpense;
                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + allOtherExpenses);
                        break;
                    default:
                }
            }
        }
    }
}



