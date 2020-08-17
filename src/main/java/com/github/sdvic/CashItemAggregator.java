package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200817
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.HashMap;

public class CashItemAggregator
{
    public int payingStudentsActual;
    int targetMonthColumnIndex;
    private int cashCombinedDirectPublicSupport;
    private int cashCombinedTuitionFees;
    private int cashCombinedSalaries;
    private int cashCombinedFacilitiesAndFacilities;
    private int cashCombinedOperations;
    private int cashOnlyIncome;
    private int cashOnlyExpenses;
    private int cashCombinedContractServices;
    private int totalGrantScholarship;
    private int budgetTotalExpenses;
    private int monthIncomeVariance;
    private int budgetDirectPublicSupport;
    private int directPublicSupportVariance;
    private int totalTuitionFees;
    private int totalWorkshopFees;
    private int tuitionFeeVariance;
    private int totalIncomeVariance;
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
    private int operationsVariance;
    private int totalExpenseVariance;
    private int payingStudentsBudget;
    private int payingStudentsVariance;
    private int budgetTuitionFees;
    private int budgetTotalIncome;
    private int budgetMonthProfit;
    private int cashOnlyProfit;
    private int pandlTotalIncome;
    private int pandlLeagueScholarship;
    private int pandlContributedServices;
    private int pandlGiftsInKindGoods;

    public void aggregateBudget(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, int targetMonthColumnIndex)
    {
        this.targetMonthColumnIndex = targetMonthColumnIndex;
        for (String key : budgetMap.keySet())
        {
            String switchKey = key.trim();
            switch (switchKey)
            {
                case "Cash Combined Direct Public Support":
                    budgetDirectPublicSupport = budgetMap.get("Cash Combined Direct Public Support");
                    totalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    pandlTotalIncome = pandLmap.get("Total Income");
                    pandlLeagueScholarship = pandLmap.get("Total 47203 League Scholarship");
                    pandlContributedServices = pandLmap.get("43460 Contributed Services");
                    pandlGiftsInKindGoods = pandLmap.get("43440 Gifts in Kind - Goods");
                    break;
                case "Cash Combined Tuition Fees":
                    totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    budgetTuitionFees = budgetMap.get("Cash Combined Tuition Fees");
                    break;
                case "Cash Only Income":
                    budgetTotalIncome = budgetMap.get("Cash Only Income");
                    break;
                case "Cash Combined Salaries":
                    totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    budgetTotalSalaries = budgetMap.get("Cash Combined Salaries");
                    break;
                case "Cash Combined Contract Services":
                    cashCombinedContractServices = pandLmap.get("Total 62100 Contract Services");
                    budgetContractServices = budgetMap.get("Cash Combined Contract Services");
                    break;
                case "Cash Combined Facilities and Equipment":
                    pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    budgetFacilities = budgetMap.get("Cash Combined Facilities and Equipment");
                    break;
                case "Cash Combined Operations":
                    supplies = pandLmap.get("Total 65040 Supplies");
                    operations = pandLmap.get("Total 65000 Operations");
                    pandlTotalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    budgetOperations = budgetMap.get("Cash Combined Operations");
                    break;
                case "Cash Only Expenses":
                    budgetTotalExpenses = budgetMap.get("Cash Only Expenses");
                    break;
                case "Cash Only Profit":
                    budgetMonthProfit = budgetMap.get("Cash Only Profit");
                    break;
                case "Paying Students (Actual)":
                    payingStudentsActual = budgetMap.get("Paying Students (Actual)");
                    break;
                case "Paying Students (Budget)":
                    payingStudentsBudget = budgetMap.get("Paying Students (Budget)");
                    break;
                default:
            }
        }
        /***************************************************************************************************************
         * Compute Combined Cash Budget Figures
         ***************************************************************************************************************/
        monthIncomeVariance = cashOnlyIncome - budgetTotalIncome;
        cashOnlyProfit = cashOnlyIncome - cashOnlyExpenses;
        cashOnlyExpenses = cashCombinedSalaries + cashCombinedContractServices + cashCombinedFacilitiesAndFacilities + cashCombinedOperations;
        totalExpenseVariance = cashOnlyExpenses - budgetTotalExpenses;
        cashCombinedOperations = supplies + operations + pandlTotalExpenses + travel + penalties;
        operationsVariance = cashCombinedOperations - budgetOperations;
        contractServiceVariance = cashCombinedContractServices - budgetContractServices;
        cashCombinedFacilitiesAndFacilities = pandlFacilitiesAndEquipment - pandlDepreciation;
        facilitiesVariance = cashCombinedFacilitiesAndFacilities - budgetFacilities;
        cashCombinedSalaries = totalSalaries + payrollServiceFees;
        monthSalaryVariance = cashCombinedSalaries - budgetTotalSalaries;
        cashOnlyIncome = cashCombinedDirectPublicSupport + cashCombinedTuitionFees + totalGrantScholarship;
        totalIncomeVariance = cashOnlyIncome - budgetTotalIncome;
        cashCombinedTuitionFees = totalTuitionFees + totalWorkshopFees;
        tuitionFeeVariance = cashCombinedTuitionFees - budgetTuitionFees;
        cashCombinedDirectPublicSupport = pandlTotalIncome - pandlLeagueScholarship - pandlContributedServices - pandlGiftsInKindGoods;
        cashOnlyIncome = cashCombinedDirectPublicSupport + cashCombinedTuitionFees + totalGrantScholarship;
        directPublicSupportVariance = cashCombinedDirectPublicSupport - budgetDirectPublicSupport;
        payingStudentsVariance = payingStudentsActual - payingStudentsBudget;
        /***************************************************************************************************************
         * Print Budget Proof Figures
         ***************************************************************************************************************/
        System.out.printf("%-40s %-20s %-20s %-20s %n", "BUDGET ACCOUNT", "BUDGET AMOUNT", "Actual AMOUNT", "Month " + targetMonthColumnIndex + " VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "----------------");
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Combined Direct Public Support", budgetDirectPublicSupport, cashCombinedDirectPublicSupport, directPublicSupportVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Combined Tuition  Fees", budgetTuitionFees, cashCombinedTuitionFees, tuitionFeeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Only Income", budgetTotalIncome, cashOnlyIncome, totalIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Combined Salaries", budgetTotalSalaries, cashCombinedSalaries, monthSalaryVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Combined Contract Services", budgetContractServices, cashCombinedContractServices, contractServiceVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Combined Facilities and Equipment", budgetFacilities, cashCombinedFacilitiesAndFacilities, facilitiesVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Combined Operations", budgetOperations, cashCombinedOperations, operationsVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Only Expenses", budgetTotalExpenses, cashOnlyExpenses, totalExpenseVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Cash Only Profit", budgetMonthProfit, cashOnlyProfit, monthIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Paying Students", payingStudentsBudget, payingStudentsActual, payingStudentsVariance);
        System.out.println("\nfinished aggregating budget with QuckBooks P&L");
        System.out.println("Finished computing budget/pandl items");

    }
    public void updateBudgetWorkbook(XSSFWorkbook budgetWorkbook, int targetMonthColumnIndex)
    {
        LocalDate localDate = LocalDate.now();
        Date date = Date.from(localDate.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());
        XSSFSheet budgetSheet = budgetWorkbook.getSheetAt(0);
        budgetSheet.getRow(0).createCell(13, XSSFCell.CELL_TYPE_STRING);
        budgetSheet.getRow(1).createCell(13, XSSFCell.CELL_TYPE_STRING);

        for (Row row : budgetSheet)
        {
            row.createCell(13, XSSFCell.CELL_TYPE_NUMERIC);//For month variance numbers
            if (row.getCell(0) != null)
            {
                switch (row.getCell(0).getStringCellValue())
                {
                    case "Cash Combined Direct Public Support":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashCombinedDirectPublicSupport);
                        row.getCell(13).setCellValue(directPublicSupportVariance);
                        break;
                    case "Cash Combined Tuition Fees":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashCombinedTuitionFees);
                        row.getCell(13).setCellValue(tuitionFeeVariance);
                        break;
                    case "Cash Only Income":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashOnlyIncome);
                        row.getCell(13).setCellValue(totalIncomeVariance);
                        break;
                    case "Cash Combined Salaries":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashCombinedSalaries);
                        row.getCell(13).setCellValue(monthSalaryVariance);
                        break;
                    case "Cash Combined Contract Services":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashCombinedContractServices);
                        row.getCell(13).setCellValue(contractServiceVariance);
                        break;
                    case "Cash Combined Facilities and Equipment":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashCombinedFacilitiesAndFacilities);
                        row.getCell(13).setCellValue(facilitiesVariance);
                        break;
                    case "Cash Combined Operations":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashCombinedOperations);
                        row.getCell(13).setCellValue(operationsVariance);
                        break;
                    case "Cash Only Expenses":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashOnlyExpenses);
                        row.getCell(13).setCellValue(totalExpenseVariance);
                        break;
                    case "Cash Only Profit":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashOnlyProfit);
                        row.getCell(13).setCellValue(monthIncomeVariance);
                        break;
                    case "Paying Students (Budget)":
                        row.getCell(13).setCellValue(payingStudentsVariance);
                        break;
                    default:
                }
            }
        }
        budgetSheet.getRow(0).getCell(0).setCellValue("Updated: " + date);
        budgetSheet.getRow(0).getCell(13).setCellValue("Month " + targetMonthColumnIndex);
        budgetSheet.getRow(1).getCell(13).setCellValue("VARIANCE");
        budgetSheet.getRow(1).getCell(targetMonthColumnIndex).setCellValue(">ACTUAL<");
        System.out.println("Finished updating budget XSSFsheet");
    }
}



