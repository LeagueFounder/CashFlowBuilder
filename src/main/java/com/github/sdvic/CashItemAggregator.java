package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200712
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.Cell;
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
    private int contributedServices;
    private int facilitiesAndEquipment;
    private int depreciation;
    private int vicFacilities;
    private int supplies;
    private int operations;
    private int totalExpenses;
    private int travel;
    private int penalties;
    private int vicOperations;
    private int investments;
    private int vicTotalIncome;
    private int vicTotalExpenses;
    private int plContractServices;
    private int netIncome;
    private Date now;
    private Calendar cals;
    private Cell currentBudgetCell;
    private Cell varianceCell;
    private int vicContractServices;
    private int budgetContractServices;
    private int contractServiceVariance;
    private int varianceColumnIndex = 13;

    public void aggregateBudget(XSSFWorkbook budgetWorkbook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkbook = budgetWorkbook;
        budgetSheet = budgetWorkbook.getSheetAt(0);
        for (Row currentBudgetRow : budgetSheet) //Iterate through budget Excel sheet
        {
            budgetSheet.getRow(0).getCell(varianceColumnIndex).setCellValue("Month " + targetMonth);
            budgetSheet.getRow(1).getCell(varianceColumnIndex).setCellValue("VARIANCE");
            budgetSheet.getRow(1).getCell(targetMonth).setCellValue("Actual");
            currentBudgetCell = currentBudgetRow.getCell(targetMonth);
            String switchKey = currentBudgetRow.getCell(0).getStringCellValue();
            switch (switchKey.trim())
            {
                case "Total 43400 Direct Public Support":
                    pandlCorporateContributions = pandLmap.get("Total 43400 Direct Public Support");
                    pandlIndividualBusinessContributions = pandLmap.get("43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("43455 Grants");
                    contributedServices = pandLmap.get("43460 Contributed Services");
                    investments = pandLmap.get("Total 45000 Investments");
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contributedServices + investments;
                    budgetSheet.getRow(currentBudgetRow.getRowNum()).getCell(targetMonth).setCellValue(vicDirectPublicSupport);
                    currentBudgetCell.setCellValue(vicDirectPublicSupport);
                    break;
                case "Total 47201 Tuition  Fees":
                    totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                    currentBudgetCell.setCellValue(vicTuitionFees);
                    break;
                case "Total Income":
                    vicTotalIncome = vicDirectPublicSupport + vicTuitionFees;
                    currentBudgetCell.setCellValue(vicTotalIncome);
                    break;
                case "Total 62000 Salaries & Related Expenses":
                    totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    vicTotalSalaries = totalSalaries + payrollServiceFees;
                    currentBudgetCell.setCellValue(vicTotalSalaries);
                    break;
                case "Total 62100 Contract Services":
                    varianceCell = currentBudgetRow.getCell(varianceColumnIndex);
                    plContractServices = pandLmap.get("Total 62100 Contract Services");
                    budgetContractServices = (int) currentBudgetCell.getNumericCellValue();
                    contractServiceVariance = plContractServices - budgetContractServices;
                    currentBudgetCell.setCellValue(plContractServices);
                    varianceCell.setCellValue(budgetContractServices - plContractServices);
                    break;
                case "Total 62800 Facilities and Equipment":
                    facilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    depreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    vicFacilities = facilitiesAndEquipment - depreciation;
                    currentBudgetCell.setCellValue(vicFacilities);
                    break;
                case "Total 65000 Operations":
                    supplies = pandLmap.get("Total 65040 Supplies");
                    operations = pandLmap.get("Total 65000 Operations");
                    totalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    vicOperations = supplies + operations + totalExpenses + travel + penalties;
                    currentBudgetCell.setCellValue(vicOperations);
                    break;
                case "Total Expenses":
                    varianceCell = currentBudgetRow.getCell(varianceColumnIndex);
                    totalExpenses = pandLmap.get("Total Expenses");
                    vicTotalExpenses = vicTotalSalaries + vicContractServices + vicFacilities + vicOperations;
                    varianceCell.setCellValue(totalExpenses - vicTotalExpenses);
                    currentBudgetCell.setCellValue(totalExpenses);
                    break;
                case "Net Income":
                    varianceCell = currentBudgetRow.getCell(varianceColumnIndex);
                    netIncome = vicTotalIncome - totalExpenses;
                    varianceCell.setCellValue(netIncome - currentBudgetCell.getNumericCellValue());
                    currentBudgetCell.setCellValue(netIncome);
                    break;
                default:
            }
        }
        cals = Calendar.getInstance();
        now = cals.getTime();
    }
    public XSSFWorkbook getBudgetWorkbook()
    {
        return budgetWorkbook;
    }
}





