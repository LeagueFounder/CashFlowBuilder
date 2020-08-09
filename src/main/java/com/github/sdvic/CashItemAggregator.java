package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200809
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.HashMap;

public class CashItemAggregator
{
    public int payingStudentsActual;
    LocalDate now;
    int targetMonthColumnIndex;
    private int vicDirectPublicSupport;
    private int vicTuitionFees;
    private int vicSalaries;
    private int vicFacilities;
    private int vicOperations;
    private int vicTotalIncome;
    private int vicExpenses;
    private int vicContractServices;
    private int totalGrantScholarship;
    private int budgetTotalExpenses;
    private int monthIncomeVariance;
    private int pandlCorporateContributions;
    private int pandlIndividualBusinessContributions;
    private int pandlGrants;
    private int contributedServices;
    private int investments;
    private int budgetDirectPublicSupport;
    private int budgetDirectPublicSupportVariance;
    private int totalTuitionFees;
    private int totalWorkshopFees;
    private int budgetTuitionFeeVariance;
    private int budgetTotalIncomeVariance;
    private int totalSalaries;
    private int budgetTotalSalaries;
    private int budgetContractServices;
    private int payrollServiceFees;
    private int monthSalaryVariance;
    private int contractServiceVariance;
    private int budgetFacilities;
    private int facilitiesVariance;
    private int pandlFacilitiesAndEquipment;
    private int pandlDepreciation;
    private int supplies;
    private int operations;
    private int pandlTotalExpenses;
    private int travel;
    private int penalties;
    private int budgetOperations;
    private int operationsVarience;
    private int totalExpenseVariance;
    private int vicNetIncome;
    private int monthBudgetIncome;
    private int payingStudentsBudget;
    private int payingStudentsVariance;
    private int budgetTuitionFees;
    private int budgetTotalIncome;
    private int monthVarianceColumnIndex = 13;
    private int ytdVarianceColumnIndex = 14;
    private int updateHeaderColumnIndex = 0;
    private int header0RowIndex = 0;
    private int header1RowIndex = 1;

    public void aggregateBudget(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, int targetMonthColumnIndex)
    {
        this.targetMonthColumnIndex = targetMonthColumnIndex;
        for (String key : budgetMap.keySet())
        {
            String switchKey = key;
            switch (switchKey)
            {
                case "Cash Combined Direct Public Support":
                    pandlCorporateContributions = pandLmap.get("Total 43400 Direct Public Support");
                    pandlIndividualBusinessContributions = pandLmap.get("43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("43455 Grants");
                    contributedServices = pandLmap.get("43460 Contributed Services");
                    investments = pandLmap.get("Total 45000 Investments");
                    budgetDirectPublicSupport = budgetMap.get("Cash Combined Direct Public Support");
                    totalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    break;
                case "Cash Combined Tuition Fees":
                    System.out.println("CashCombiedTuitionFees");
                    totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    budgetTuitionFees = budgetMap.get("Cash Combined Tuition Fees");
                    break;
                case "Total Cash Only Income":
                    budgetTotalIncome = budgetMap.get("Total Cash Only Income");
                    budgetTotalIncomeVariance = (int) (vicTotalIncome - budgetTotalIncome);
                    break;
                case "Cash Combined Salaries":
                    totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    budgetTotalSalaries = budgetMap.get("Cash Combined Salaries");
//                    monthVarianceCell.setCellValue(monthSalaryVariance);
//                    currentBudgetCell.setCellValue(vicSalaries);
                    break;
                case "Cash Combined Contract Services":
                    vicContractServices = pandLmap.get("Total 62100 Contract Services");
                    budgetContractServices = budgetMap.get("Cash Combined Contract Services");
//                    monthVarianceCell.setCellValue(contractServiceVariance);
//                    currentBudgetCell.setCellValue(vicContractServices);
                    break;
                case "Cash Combined Facilities and Equipment":
                    pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    budgetFacilities = budgetMap.get("Cash Combined Facilities and Equipment");
//                    monthVarianceCell.setCellValue(facilitiesVariance);
//                    currentBudgetCell.setCellValue(vicFacilities);
                    break;
                case "Cash Combined Operations":
                    supplies = pandLmap.get("Total 65040 Supplies");
                    operations = pandLmap.get("Total 65000 Operations");
                    pandlTotalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    budgetOperations = budgetMap.get("Cash Combined Operations");
                    operationsVarience = (int) (vicOperations - budgetOperations);
//                    monthVarianceCell.setCellValue(operationsVarience);
//                    currentBudgetCell.setCellValue(vicOperations);
                    break;
                case "Total Cash Only Expenses":
                    budgetTotalExpenses = budgetMap.get("Total Cash Only Expenses");
                    //budgetTotalExpenses = (int) currentBudgetCell.getNumericCellValue();
//                    monthVarianceCell.setCellValue(totalExpenseVariance);
//                    currentBudgetCell.setCellValue(vicExpenses);
                    break;
                case "Net Cash Only Profit":
                    monthBudgetIncome = budgetMap.get("Net Cash Only Profit");
//                    monthVarianceCell.setCellValue(monthIncomeVariance);
//                    currentBudgetCell.setCellValue(vicNetIncome);
                    break;
                case "Net Cash Only Income VARIANCE":
//                    currentBudgetCell.setCellValue(monthIncomeVariance);
                    break;
                case "Paying Students (Actual)":
                    payingStudentsActual = budgetMap.get("Paying Students (Actual)");
                    break;
                case "Paying Students (Budget)":
                    payingStudentsBudget = budgetMap.get("Paying Students (Budget)");
//                    monthVarianceCell.setCellValue(payingStudentsVariance);
                    break;
                default:
            }
        }
        System.out.println("finished aggregating budget with QuckBooks P&L");
    }

    public void computeLineItems()
    {
        System.out.printf("%-40s %-20s %-20s %-20s %n", "BUDGET ACCOUNT", "BUDGET AMOUNT", "Actual AMOUNT", "Month " + targetMonthColumnIndex + " VARIANCE");
        /********************************************************************************
         * Direct Public Support
         ********************************************************************************/
        vicTotalIncome = vicDirectPublicSupport + vicTuitionFees + totalGrantScholarship;
        vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contributedServices + investments + totalGrantScholarship;
        budgetDirectPublicSupportVariance = vicDirectPublicSupport - budgetDirectPublicSupport;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Direct Public Support", budgetDirectPublicSupport, vicDirectPublicSupport, budgetDirectPublicSupportVariance);
        /********************************************************************************
         * Tuition
         ********************************************************************************/
        vicTuitionFees = totalTuitionFees + totalWorkshopFees;
        budgetTuitionFeeVariance = vicTuitionFees - budgetTuitionFees;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition  Fees", budgetTuitionFees, vicTuitionFees, budgetTuitionFeeVariance);
        /********************************************************************************
         * Total Income
         ********************************************************************************/
        vicTotalIncome = vicDirectPublicSupport + vicTuitionFees + totalGrantScholarship;
        budgetTotalIncomeVariance = vicTotalIncome - budgetTotalIncome;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", (int) budgetTotalIncome, vicTotalIncome, budgetTotalIncomeVariance);
        /********************************************************************************
         * Salaries
         ********************************************************************************/
        vicSalaries = totalSalaries + payrollServiceFees;
        monthSalaryVariance = vicSalaries - budgetTotalSalaries;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", (int) budgetTotalSalaries, vicSalaries, monthSalaryVariance);
        /********************************************************************************
         * Contract Services
         ********************************************************************************/
        contractServiceVariance = vicContractServices - budgetContractServices;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Contract Services", (int) budgetContractServices, vicContractServices, contractServiceVariance);
        /********************************************************************************
         * Facilities
         ********************************************************************************/
        vicFacilities = pandlFacilitiesAndEquipment - pandlDepreciation;
        facilitiesVariance = vicFacilities - budgetFacilities;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Facilities and Equipment", (int) budgetFacilities, vicFacilities, facilitiesVariance);
        /********************************************************************************
         * Operations
         ********************************************************************************/
        vicOperations = supplies + operations + pandlTotalExpenses + travel + penalties;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", (int) budgetOperations, vicOperations, vicOperations - (int) budgetOperations);
        /********************************************************************************
         * Total Expenses
         ********************************************************************************/
        vicExpenses = vicSalaries + vicContractServices + vicFacilities + vicOperations;
        totalExpenseVariance = vicExpenses - budgetTotalExpenses;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Expenses", (int) budgetTotalExpenses, vicExpenses, totalExpenseVariance);
        /********************************************************************************
         * Net Income
         ********************************************************************************/
        vicNetIncome = vicTotalIncome - vicExpenses;
        monthIncomeVariance = vicNetIncome - monthBudgetIncome;
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "NetCashOnlyIncom", (int) budgetTotalIncome, vicNetIncome, monthIncomeVariance);
        /********************************************************************************
         * Students
         ********************************************************************************/
        payingStudentsVariance = payingStudentsActual - payingStudentsBudget;
        System.out.printf("%-40s %,-40d %n", "Paying Students (Actual)", payingStudentsActual);
        System.out.printf("%-40s %,-40d %n", "Paying Students (Budget)", payingStudentsBudget);
        System.out.println("Finished computing budget/pandl itms");
    }

    public void updateBudgetWorkbook(XSSFWorkbook budgetWorkbook, int targetMonthColumnIndex)
    {
        LocalDateTime now = LocalDateTime.now();
        XSSFSheet budgetSheet = budgetWorkbook.getSheetAt(0);
        budgetSheet.getRow(0).createCell(13, XSSFCell.CELL_TYPE_STRING);
        budgetSheet.getRow(0).createCell(14, XSSFCell.CELL_TYPE_STRING);
        budgetSheet.getRow(1).createCell(13, XSSFCell.CELL_TYPE_STRING);
        budgetSheet.getRow(1).createCell(14, XSSFCell.CELL_TYPE_STRING);
        for (Row row: budgetSheet)
        {
            row.createCell(13, XSSFCell.CELL_TYPE_NUMERIC);
        }
        budgetSheet.getRow(4).createCell(13, XSSFCell.CELL_TYPE_NUMERIC);//Month Tuition variance
        budgetSheet.getRow(0).getCell(0).setCellValue("Updated: " + now);
        budgetSheet.getRow(0).getCell(13).setCellValue("Month " + targetMonthColumnIndex);
        budgetSheet.getRow(1).getCell(13).setCellValue("VARIANCE");
        budgetSheet.getRow(0).getCell(14).setCellValue("YTD");
        budgetSheet.getRow(1).getCell(14).setCellValue("VARIANCE");
        budgetSheet.getRow(4).getCell(13).setCellValue(budgetTuitionFeeVariance);
        budgetSheet.getRow(3).getCell(13).setCellValue(budgetDirectPublicSupportVariance);
        System.out.println("Aggregating, Month " + targetMonthColumnIndex + " Budget Proof => ");
        System.out.println("Finished updating budget XSSFsheet");
    }
}



