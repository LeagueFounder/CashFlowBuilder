package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190504
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.util.HashMap;

public class CashFlowItemAggregator
{
    public int cashContributions;
    int cashTuition;
    int payrollExpense;
    int facilityExpense;
    int operationsExpense;
    int otherExpense;
    int travelExpense;
    int totalOtherExpense;
    int cashGrants;
    int workshopFees;
    int year = 6;

    public CashFlowItemAggregator(Workbook sarah5yearWorkbook, HashMap<String, Integer> pandLmap, String version)
    {
        FormulaEvaluator evaluator = sarah5yearWorkbook.getCreationHelper().createFormulaEvaluator();

        Sheet fiveYearSheet = sarah5yearWorkbook.getSheetAt(0);
        Cell revDateCell = fiveYearSheet.getRow(1).getCell(0);
        revDateCell.setCellValue(version);
        for (Row row : fiveYearSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            Cell projectionLabelCell = row.getCell(0);
            if (projectionLabelCell.getCellType() == Cell.CELL_TYPE_STRING)
            {
                switch (projectionLabelCell.getStringCellValue())
                {
                    case "Cash Grants":
                        Cell cashGrantsCell = fiveYearSheet.getRow(4).getCell(year);
                        cashGrants = pandLmap.get("      Total 47204 Grant Scholarship");
                        cashGrantsCell.setCellValue(cashGrants);
                        System.out.println(projectionLabelCell.getStringCellValue() + " => " + cashGrants);
                        break;
                    case "Cash Contributions":
                        Cell cashContributionsCell = fiveYearSheet.getRow(3).getCell(year);
                        cashContributions = pandLmap.get("      43450 Individ, Business Contributions");
                        cashContributionsCell.setCellValue(cashContributions);
                        System.out.println(projectionLabelCell.getStringCellValue() + " => " + cashContributions);
                        break;
                    case "Cash Tuition ":
                        Cell cashTuitionCell = fiveYearSheet.getRow(5).getCell(year);
                        cashTuition = pandLmap.get("      Total 47201 Tuition  Fees");
                        workshopFees = pandLmap.get("      Total 47202 Workshop Fees");
                        cashTuitionCell.setCellValue((int)(cashTuition + workshopFees));
                        System.out.println(projectionLabelCell.getStringCellValue() + " => " + (cashTuition + workshopFees));
                        break;
                    case "Payroll":
                        Cell payrollCell = fiveYearSheet.getRow(9).getCell(year);
                        payrollExpense = pandLmap.get("   Total 62000 Salaries & Related Expenses");
                        payrollCell.setCellValue(payrollExpense);
                        System.out.println(projectionLabelCell.getStringCellValue() + " => " + payrollExpense);
                        break;
                    case "Facilities":
                        Cell facilitiesCell = fiveYearSheet.getRow(10).getCell(year);
                        facilityExpense = pandLmap.get("   Total 62800 Facilities and Equipment");
                        facilitiesCell.setCellValue(facilityExpense);
                        System.out.println(projectionLabelCell.getStringCellValue() + " => " + facilitiesCell.getNumericCellValue());
                        break;
                    case "All Other Expenses":
                        Cell otherExpenseCell = fiveYearSheet.getRow(11).getCell(year);
                        operationsExpense = pandLmap.get("   Total 65000 Operations");
                        otherExpense = pandLmap.get("   Total 65100 Other Types of Expenses");
                        travelExpense = pandLmap.get("   Total 68300 Travel and Meetings");
                        otherExpenseCell.setCellValue(operationsExpense + otherExpense + travelExpense + totalOtherExpense);
                        System.out.println(projectionLabelCell.getStringCellValue() + " => " + otherExpenseCell.getNumericCellValue());
                        break;
                    default:
                }
            }
        }
        evaluator.evaluateAll();
    }
}



