package com.github.sdvic;
/******************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 190508
 * copyright 2019 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;

public class CashItemAggregator
{
    private XSSFSheet budgetSheet;
    private XSSFWorkbook budgetWorkBook;
    private Cell accountItemCell;
    private Cell directPublicSupport;//Budget item
    private int pandlValue;
    private int pandlCorporateContributions;
    private int pandlIndividualBusinessContributions;
    private int pandlGrants;
    private int vicDirectPublicSupport;
    private int totalTuitionFees;
    private int totalWorkshopFees;
    private int vicTuitionFees;
    private int totalSalaries;
    private int payrollServiceFees;
    private int vicTotalSalaries;
    private int contractServices;
    private int facilitiesAndEquipment;
    private int depreciation;
    private int vicFacilities;

    public CashItemAggregator(XSSFWorkbook budgetWorkBook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkBook = budgetWorkBook;
        budgetSheet = budgetWorkBook.getSheetAt(0);//Go through budget looking at column targetMonth
        for (Row row : budgetSheet) //Bring in budget sheet
        {
            accountItemCell = row.getCell(0);
            switch (accountItemCell.getStringCellValue())
            {
                case "   Total 43400 Direct Public Support":
                    pandlCorporateContributions = pandLmap.get("      43410 Corporate Contributions");
                    pandlIndividualBusinessContributions = pandLmap.get("      43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("      43455 Grants");
                    System.out.println("pandl corp cont = " + pandlCorporateContributions);
                    System.out.println("43450 Individ, Business Contributions = " + pandlIndividualBusinessContributions);
                    System.out.println("43455 Grants = " + pandlGrants);
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants;
                    System.out.println("vic direct public support = " + vicDirectPublicSupport);
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicDirectPublicSupport);
                    System.out.println("changed dps = " + budgetSheet.getRow(row.getRowNum()).getCell(targetMonth));
                    break;
                    case "      Total 47201 Tuition  Fees":
                        totalTuitionFees = pandLmap.get("      Total 47201 Tuition  Fees");
                        totalWorkshopFees = pandLmap.get("      Total 47202 Workshop Fees");
                        vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                        budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicTuitionFees);
                        break;
                case "   Total 62000 Salaries & Related Expenses":
                    totalSalaries = pandLmap.get("   Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("   62145 Payroll Service Fees");
                    vicTotalSalaries = totalSalaries + payrollServiceFees;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicTotalSalaries);
                    break;
                case "   Total 62100 Contract Services":
                    contractServices = pandLmap.get("   Total 62100 Contract Services");
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(contractServices);
                    break;

                case "   Total 62800 Facilities and Equipment":
                    facilitiesAndEquipment = pandLmap.get("   Total 62800 Facilities and Equipment");
                    depreciation = pandLmap.get("      62810 Depr and Amort - Allowable");
                    vicFacilities = facilitiesAndEquipment - depreciation;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicFacilities);
                    break;
//                    case "All Other Expenses":
//                        Cell otherExpenseCell = fiveYearSheet.getRow(11).getCell(2);
//                        otherExpenseCell.setCellType(Cell.CELL_TYPE_NUMERIC);
//                        int operationsExpense = pandLmap.get("   Total 65000 Operations");
//                        int otherExpense = pandLmap.get("   Total 65100 Other Types of Expenses");
//                        int travelExpense = pandLmap.get("   Total 68300 Travel and Meetings");
//                        int contractServiceExpense = 0;
//                        int allOtherExpenses = 0;
//                        if (pandLmap.get("   Total 62100 Contract Services") != null)
//                        {
//                            otherExpenseCell.setCellValue(allOtherExpenses);
//                        }
//                        allOtherExpenses = operationsExpense + otherExpense + travelExpense + contractServiceExpense;
//                        System.out.println(fiveYearProjectionLabelCell.getStringCellValue() + " => " + allOtherExpenses);
//                        break;
                default:
            }
//            for (Row row2 : budgetSheet)
//            {
//                for (Cell cell2 : row2)
//                {
//                    System.out.println(cell2);
//                }
//            }
        }
    }
}



