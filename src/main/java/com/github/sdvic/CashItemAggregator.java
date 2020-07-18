package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200717
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

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
    private int pandlFacilitiesAndEquipment;
    private int pandlDepreciation;
    private int vicFacilities;
    private int supplies;
    private int operations;
    private int pandlTotalExpenses;
    private int travel;
    private int penalties;
    private int vicOperations;
    private int investments;
    private int vicTotalIncome;
    private int vicTotalExpenses;
    private int plContractServices;
    private int vicNetIncome;
    private Cell currentBudgetCell;
    private Cell varianceCell;
    private int vicContractServices;
    private int budgetContractServices;
    private int contractServiceVariance;
    private int monthVarianceColumnIndex = 13;
    private int totalGrantScholarship;
    private int budgetDirectPublicSupport;
    private int budgetDirectPublicSupportVariance;
    private int pandlDirectPublicSupport;
    private int budgetTuitionFeeVariance;
    private double budgetTotalIncome;
    private int budgetTotalIncomeVariance;
    private int budgetTotalSalaries;
    private int budgetOperationsValue;
    private int operationsVarience;
    private double budgetFacilities;
    private double payingStudents;
    private double payingStudentsBudget;
    private int payingStudentsDerived;
    private double payingStudentsVariation;

    public void aggregateBudget(XSSFWorkbook budgetWorkbook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkbook = budgetWorkbook;
        budgetSheet = budgetWorkbook.getSheetAt(0);
        System.out.println("Aggregating, Month " + targetMonth + " Resultant Budget Proof, budget sheet numbers:");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "BUDGET ACCOUNT", "BUDGET AMOUNT", "PandL AMOUNT", "VARIANCE");
        for (Row currentBudgetRow : budgetSheet) //Iterate through budget Excel sheet
        {
            budgetSheet.getRow(0).getCell(monthVarianceColumnIndex).setCellValue("Month " + targetMonth);
            budgetSheet.getRow(1).getCell(targetMonth).setCellValue("Actual");
            currentBudgetCell = currentBudgetRow.getCell(targetMonth);
            varianceCell = currentBudgetRow.getCell(monthVarianceColumnIndex);
            String switchKey = currentBudgetRow.getCell(0).getStringCellValue();
            switch (switchKey.trim())
            {
                case "Total 43400 Direct Public Support":
                    pandlCorporateContributions = pandLmap.get("Total 43400 Direct Public Support");
                    pandlIndividualBusinessContributions = pandLmap.get("43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("43455 Grants");
                    contributedServices = pandLmap.get("43460 Contributed Services");
                    investments = pandLmap.get("Total 45000 Investments");
                    totalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contributedServices + investments + totalGrantScholarship;
                    budgetDirectPublicSupport = (int) currentBudgetCell.getNumericCellValue();
                    budgetDirectPublicSupportVariance = (int) vicDirectPublicSupport - budgetDirectPublicSupport;
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Direct Public Support", (int) budgetDirectPublicSupport, (int) vicDirectPublicSupport, budgetDirectPublicSupportVariance);
                    currentBudgetCell.setCellValue(vicDirectPublicSupport);
                    varianceCell.setCellValue(budgetDirectPublicSupportVariance);
                    currentBudgetRow.getCell(14).setCellValue(computeYTDvariance(currentBudgetRow));
                    break;
                case "Total 47201 Tuition  Fees":
                    totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                    budgetTuitionFeeVariance = (int) ((int) vicTuitionFees - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Tuition  Fees", (int) currentBudgetCell.getNumericCellValue(), (int) vicTuitionFees, budgetTuitionFeeVariance);
                    currentBudgetCell.setCellValue(vicTuitionFees);
                    varianceCell.setCellValue(budgetTuitionFeeVariance);
                    break;
                case "Total Income":
                    vicTotalIncome = vicDirectPublicSupport + vicTuitionFees + totalGrantScholarship;
                    budgetTotalIncome = currentBudgetCell.getNumericCellValue();
                    budgetTotalIncomeVariance = (int) (vicTotalIncome - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Total Income", (int) currentBudgetCell.getNumericCellValue(), (int) vicTotalIncome, budgetTotalIncomeVariance);
                    currentBudgetCell.setCellValue(vicTotalIncome);
                    varianceCell.setCellValue(budgetTotalIncomeVariance);
                    break;
                case "Total 62000 Salaries & Related Expenses":
                    totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    vicTotalSalaries = totalSalaries + payrollServiceFees;
                    budgetTotalSalaries = (int) currentBudgetCell.getNumericCellValue();
                    varianceCell.setCellValue(vicTotalSalaries - budgetTotalSalaries);
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Salaries", (int) currentBudgetCell.getNumericCellValue(), (int) vicTotalSalaries, vicTotalSalaries - totalSalaries);
                    currentBudgetCell.setCellValue(vicTotalSalaries);
                    varianceCell.setCellValue(vicTotalSalaries);
                    break;
                case "Total 62100 Contract Services":
                    plContractServices = pandLmap.get("Total 62100 Contract Services");
                    budgetContractServices = (int) currentBudgetCell.getNumericCellValue();
                    contractServiceVariance = plContractServices - budgetContractServices;
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Total Contract Services", (int) currentBudgetCell.getNumericCellValue(), plContractServices, contractServiceVariance);
                    currentBudgetCell.setCellValue(plContractServices);
                    varianceCell.setCellValue(contractServiceVariance);
                    break;
                case "Total 62800 Facilities and Equipment":
                    pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    vicFacilities = pandlFacilitiesAndEquipment - pandlDepreciation;
                    budgetFacilities = currentBudgetCell.getNumericCellValue();
                    varianceCell.setCellValue(vicFacilities - pandlFacilitiesAndEquipment);
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Facilities and Equipment", (int) currentBudgetCell.getNumericCellValue(), plContractServices, vicFacilities - pandlFacilitiesAndEquipment);
                    currentBudgetCell.setCellValue(vicFacilities);
                    varianceCell.setCellValue((int) vicFacilities);
                    break;
                case "Total 65000 Operations":
                    supplies = pandLmap.get("Total 65040 Supplies");
                    operations = pandLmap.get("Total 65000 Operations");
                    pandlTotalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    vicOperations = supplies + operations + pandlTotalExpenses + travel + penalties;
                    budgetOperationsValue = (int) currentBudgetCell.getNumericCellValue();
                    operationsVarience = (int) (vicOperations - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Operations", (int) currentBudgetCell.getNumericCellValue(), (int) vicOperations, (int) vicOperations - (int) currentBudgetCell.getNumericCellValue());
                    currentBudgetCell.setCellValue(vicOperations);
                    varianceCell.setCellValue(operationsVarience);
                    break;
                case "Total Expenses":
                    varianceCell = currentBudgetRow.getCell(monthVarianceColumnIndex);
                    pandlTotalExpenses = pandLmap.get("Total Expenses");
                    vicTotalExpenses = vicTotalSalaries + vicContractServices + vicFacilities + vicOperations;
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Total Expenses", (int) currentBudgetCell.getNumericCellValue(), (int) pandlTotalExpenses, (int) pandlTotalExpenses - (int) currentBudgetCell.getNumericCellValue());
                    currentBudgetCell.setCellValue(pandlTotalExpenses);
                    varianceCell.setCellValue(pandlTotalExpenses - vicTotalExpenses);
                    break;
                case "Net Cash Income":
                    vicNetIncome = vicTotalIncome - pandlTotalExpenses;
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Net Cash Income", (int) currentBudgetCell.getNumericCellValue(), (int) pandlTotalExpenses, (int) pandlTotalExpenses - (int) currentBudgetCell.getNumericCellValue());
                    currentBudgetCell.setCellValue(vicNetIncome);
                    varianceCell.setCellValue(vicNetIncome - currentBudgetCell.getNumericCellValue());
                    break;
                case "Paying Students (Budget)":
                    payingStudentsBudget = currentBudgetCell.getNumericCellValue();
                    currentBudgetCell.setCellValue(payingStudentsBudget);
                    System.out.printf("%-40s %-20d %n", "Paying Students (Budget)", (int) currentBudgetCell.getNumericCellValue());
                    break;
                case "Paying Students (Derived)":
                    payingStudentsDerived = totalTuitionFees / 240;
                    currentBudgetCell.setCellValue(payingStudentsDerived);
                    System.out.printf("%-40s %-20d %n", "Paying Students (Derived)", (int) currentBudgetCell.getNumericCellValue());
                    break;
                case "Paying Students VARIATION":
                    payingStudentsVariation = payingStudentsDerived - payingStudentsBudget;
                    currentBudgetCell.setCellValue(payingStudentsVariation);
                    System.out.printf("%-40s %-20d %n", "Paying Students VARIATION", (int) currentBudgetCell.getNumericCellValue());
                    break;
                default:
            }
        }
    }

    public XSSFWorkbook getBudgetWorkbook()
    {
        return budgetWorkbook;
    }

    public int computeYTDvariance(Row budgetItemRow)
    {
        int ytdVariance = 0;
        int i = 1;
        while(budgetSheet.getRow(1).getCell(i).getStringCellValue().equals("Actual"))
        {
            if (budgetItemRow.getCell(i).getCellType() == CELL_TYPE_NUMERIC)
            {
                ytdVariance += budgetItemRow.getCell(i).getNumericCellValue();
            }
            i++;
        }
        return ytdVariance;
    }

}



