package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200709
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;

public class CashItemAggregator
{
    private XSSFWorkbook budgetWorkbook;
    private XSSFSheet budgetSheet;
    private Cell accountItemCell;
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
    private int supplies;
    private int operations;
    private int expenses;
    private int travel;
    private int penalties;
    private int vicOperations;
    private int contributedServices;
    private HashMap<String, Integer> pandLmap;
    private Date now;
    private Calendar cals;
    public void aggregateBudget(XSSFWorkbook budgetWorkbook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkbook = budgetWorkbook;
        budgetSheet = budgetWorkbook.getSheetAt(0);
        //budgetSheet.getRow(0).getCell(0).setCellValue(now);
        for (Row row : budgetSheet) //Iterate through budget Excel sheet
        {
            accountItemCell = row.getCell(0);
            switch (accountItemCell.getStringCellValue())
            {
                case "   Total 43400 Direct Public Support":
                    System.out.println("***************found: Total 43400 Direct Public Support");
                    pandlCorporateContributions = pandLmap.get("   Total 43400 Direct Public Support");
                    pandlIndividualBusinessContributions = pandLmap.get("      43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("      43455 Grants");
                    contractServices = pandLmap.get("      43460 Contributed Services");
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contractServices;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicDirectPublicSupport);
                    System.out.println(vicDirectPublicSupport);
                    break;
                case "      Total 47201 Tuition  Fees":
                    System.out.println("*****************found: Total 47201 Tuition  Fees");
                    totalTuitionFees = pandLmap.get("      Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("      Total 47202 Workshop Fees");
                    vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicTuitionFees);
                    System.out.println(vicTuitionFees);
                    break;
                case "   Total 62000 Salaries & Related Expenses":
                    System.out.println("****************found: Total 62000 Salaries & Related Expenses");
                    totalSalaries = pandLmap.get("   Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("   62145 Payroll Service Fees");
                    vicTotalSalaries = totalSalaries + payrollServiceFees;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicTotalSalaries);
                    System.out.println(vicTotalSalaries);
                    break;
                case "   Total 62100 Contract Services":
                    System.out.println("*******************found: Total 62100 Contract Services ");
                    contractServices = pandLmap.get("   Total 62100 Contract Services");
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(contractServices);
                    System.out.println(contractServices);
                    break;

                case "   Total 62800 Facilities and Equipment":
                    System.out.println("*************found: Total 62800 Facilities and Equipment");
                    facilitiesAndEquipment = pandLmap.get("   Total 62800 Facilities and Equipment");
                    depreciation = pandLmap.get("      62810 Depr and Amort - Allowable");
                    vicFacilities = facilitiesAndEquipment - depreciation;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicFacilities);
                    System.out.println(vicFacilities);
                    break;
                case "   Total 65000 Operations":
                    System.out.println("************found: Total 65000 Operations");
                    supplies = pandLmap.get("      Total 65040 Supplies");
                    operations = pandLmap.get("   Total 65000 Operations");
                    expenses = pandLmap.get("   Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("   Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("   90100 Penalties");
                    vicOperations = supplies + operations + expenses + travel + penalties;
                    budgetSheet.getRow(row.getRowNum()).getCell(targetMonth).setCellValue(vicOperations);
                    System.out.println(vicOperations);
                    break;
                default:
            }
        }
        cals = Calendar.getInstance();
        now = cals.getTime();
        budgetSheet.getRow(0).getCell(0).setCellValue("month " + targetMonth);
    }
    public XSSFWorkbook getBudgetWorkbook()
    {
        return budgetWorkbook;
    }
}




