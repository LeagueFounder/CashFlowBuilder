package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200820
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
    private int cashCombinedOperations;
    private int cashOnlyIncome;
    private int budgetCashOnlyExpenses;
    private int actualCashCombinedContractServices;
    private int pandlGrantScholarship;
    private int budgetTotalExpenses;
    private int monthIncomeVariance;
    private int budgetDirectPublicSupport;
    private int directPublicSupportVariance;
    private int pandlTuitionFees;
    private int pandlWorkshopFees;
    private int tuitionFeeVariance;
    private int totalIncomeVariance;
    private int pandlSalaries;
    private int budgetTotalSalaries;
    private int budgetContractServices;
    private int pandlPayrollServiceFees;
    private int monthSalaryVariance;
    private int contractServiceVariance;
    private int budgetFacilities;
    private int facilitiesVariance;
    private int pandlFacilitiesAndEquipment;
    private int pandlDepreciation;
    private int pandlSupplies;
    private int pandlOperations;
    private int pandlTotalExpenses;
    private int pandlTravel;
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
    private int pandlContributedServices;
    private int pandlGiftsInKindGoods;
    private int pandlLeagueScholarship;
    private int cashMiscIncome;
    private int actualCashMiscExpenses;
    private int cashCombinedFacilitiesEquipment;
    private int pandlProgramIncome;
    private int actualCashCombinedDirectPublicSupport;
    private int totalPandlDirectPublicSupport;
    private int budgetCashCombinedTuitionFees;
    private int budgetMiscIncome;
    private int pandlInvestments;
    private int pandlOtherIncome;
    private int budgetInvestments;
    private int budgetCashMiscIncome;
    private int budgetCashOnlyIncome;
    private int budgetCashTuitionFees;
    private int budgetCashMiscExpenses;
    private int budgetCashCombinedOperations;
    private int budgetCashCombinedSalaries;
    private int budgetCashCombinedContractServices;
    private int budgetCashCombinedFacilitiesEquiment;
    private int budgetCashCombinedFacilitiesEquipment;
    private int cashOnlyExpenses;
    private int totalPandlIndirectPublicSupport;
    private int pandlTotalGrantScholarship;
    private int actualCashCombinedSalaries;
    private int actualCashCombinedOperations;
    private int actualCashMiscIncome;
    private int actualTotalExpenses;
    private int actualOtherExpenses;

    /******************************************************************************************
     * Extract Budget Items from P & L Map
     ******************************************************************************************/
    public void extractBudgetItemsFromPandLMap(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetHashMap)
    {
        System.out.println("(5) Extracting Budget Items From PandL Map");
        for (String key : budgetHashMap.keySet())
        {
            String switchKey = key.trim();
            switch (switchKey)
            {
                case "Grants and Gifts":
                    totalPandlDirectPublicSupport = pandLmap.get("Total 43400 Direct Public Support");
                    //totalPandlIndirectPublicSupport = pandLmap.get("Total 44800 Indirect Public Support");
                    pandlContributedServices = pandLmap.get("43460 Contributed Services");//Non cash item...must be subtraccted
                    //pandlGiftsInKindGoods = pandLmap.get("43440 Gifts in Kind - Goods");//Non cash item...must be subtraccted
                    pandlTotalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    break;
                case "Tuition":
                    pandlProgramIncome = pandLmap.get("Total 47200 Program Income");
                    pandlTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    pandlGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    pandlWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    //pandlLeagueScholarship = pandLmap.get("Total 65090 Scholarships");//non cash item
                    break;
                case "Total Income":
                    pandlTotalIncome = pandLmap.get("Total Income");//********** major income item
                    pandlInvestments = pandLmap.get("Total 45000 Investments");
                    //pandlOtherIncome = pandLmap.get("Total 46400 Other Types of Income");
                    break;
                case "Salaries":
                    pandlSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    pandlPayrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    pandlContributedServices = pandLmap.get("62010 Salaries contributed services");//non-cash item
                    break;
                case "Contract Services":
                    actualCashCombinedContractServices = pandLmap.get("Total 62100 Contract Services");
                    break;
                case "Rent":
                    pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");//non cash item
                    break;
                case "Operations":
                    pandlSupplies = pandLmap.get("Total 65040 Supplies");
                    pandlOperations = pandLmap.get("Total 65000 Operations");
                    pandlTravel= pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    break;
                case "TotalExpenses":
                    pandlTotalExpenses = pandLmap.get("Total Expenses");//******* major expense item
                    actualOtherExpenses = pandLmap.get("Total Other Expenses");
                    break;
                case "Profit":
                    break;
                case "Paying Students (Actual)":
                    break;
                case "Paying Students (Budget)":
                    break;
                default:
            }
        }
        System.out.println("(6) Finished Extracting Budget Items From PandL Map");
    }
    /******************************************************************************************
     * Extract Budget Items from Budget Map
     ******************************************************************************************/
    public void extractBudgetItemsFromBudgetMap(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, int targetMonthColumnIndex)
    {
        System.out.println("(7) Extracting Budget Items From Budget Map");
        this.targetMonthColumnIndex = targetMonthColumnIndex;
        for (String key : budgetMap.keySet())
        {
            String switchKey = key.trim();
            switch (switchKey)
            {
                case "Grants and Gifts":
                    budgetDirectPublicSupport = budgetMap.get("Grants and Gifts");
                    break;
                case "Tuition":
                    budgetTuitionFees = budgetMap.get("Tuition");
                    break;
                case "Total Income":
                    budgetTotalIncome = budgetMap.get("Total Income");
                    break;
                case "Salaries":
                    budgetTotalSalaries = budgetMap.get("Salaries");
                    break;
                case "Contract Services":
                    budgetContractServices = budgetMap.get("Contract Services");
                    break;
                case "Rent":
                    budgetFacilities = budgetMap.get("Rent");
                    break;
                case "Operations":
                    budgetOperations = budgetMap.get("Operations");
                    break;
                case "Cash Expenses":
                    budgetTotalExpenses = budgetMap.get("Total Expenses");
                    break;
                case "Profit":
                    budgetMonthProfit = budgetMap.get("Profit");
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
        System.out.println("(8) Finished Extracting Budget Items From Budget Map");
    }
    /******************************************************************************************
     * Compute budget sheet entries
     ******************************************************************************************/
    public void computeCombinedCashBudgetSheetEntries()
    {
        System.out.println("(9) Computing Combined Budget Sheet Entries");
        actualCashCombinedDirectPublicSupport = totalPandlDirectPublicSupport + totalPandlIndirectPublicSupport + pandlGrantScholarship - pandlContributedServices - pandlGiftsInKindGoods;
        monthIncomeVariance = cashOnlyIncome - budgetTotalIncome;
        totalExpenseVariance = budgetCashOnlyExpenses - budgetTotalExpenses;
        actualCashCombinedOperations = pandlSupplies + pandlOperations + pandlTotalExpenses + pandlTravel + penalties;
        operationsVariance = cashCombinedOperations - budgetOperations;
        contractServiceVariance = actualCashCombinedContractServices - budgetContractServices;
        budgetCashCombinedFacilitiesEquiment = pandlFacilitiesAndEquipment - pandlDepreciation;
        facilitiesVariance = budgetCashCombinedFacilitiesEquiment - budgetFacilities;
        actualCashCombinedSalaries = pandlSalaries + pandlPayrollServiceFees;
        monthSalaryVariance = cashCombinedSalaries - budgetTotalSalaries;
        cashOnlyIncome = pandlTotalIncome - pandlLeagueScholarship - pandlContributedServices - pandlGiftsInKindGoods;
        cashOnlyProfit = cashOnlyIncome - budgetCashOnlyExpenses;
        totalIncomeVariance = cashOnlyIncome - budgetTotalIncome;
        tuitionFeeVariance = cashCombinedTuitionFees - budgetTuitionFees;
        directPublicSupportVariance = cashCombinedDirectPublicSupport - budgetDirectPublicSupport;
        payingStudentsVariance = payingStudentsActual - payingStudentsBudget;
        actualCashMiscIncome = pandlInvestments + pandlOtherIncome;
        budgetCashCombinedTuitionFees = pandlProgramIncome - pandlLeagueScholarship;
        budgetCashOnlyIncome = actualCashCombinedDirectPublicSupport + budgetCashCombinedTuitionFees + budgetCashMiscIncome;
        actualTotalExpenses = pandlTotalExpenses - pandlContributedServices - pandlLeagueScholarship + actualOtherExpenses;
        budgetCashOnlyExpenses = budgetTotalSalaries + budgetCashCombinedFacilitiesEquipment + budgetCashCombinedFacilitiesEquipment + budgetCashCombinedOperations + budgetCashMiscExpenses;
        /***************************************************************************************************************
         * Print Budget Proof Figures
         ***************************************************************************************************************/
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "BUDGET ACCOUNT", "BUDGET AMOUNT", "Actual AMOUNT", "Month " + targetMonthColumnIndex + " VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "----------------");
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Grants and Gifts", budgetDirectPublicSupport, actualCashCombinedDirectPublicSupport, directPublicSupportVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition", budgetTuitionFees, budgetCashCombinedTuitionFees, tuitionFeeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", budgetTotalIncome, budgetCashOnlyIncome, totalIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", budgetTotalSalaries, actualCashCombinedSalaries, monthSalaryVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Contract Services", budgetContractServices, actualCashCombinedContractServices, contractServiceVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Rent",  budgetFacilities, budgetCashCombinedFacilitiesEquiment, facilitiesVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", budgetOperations, actualCashCombinedOperations, operationsVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "TotalExpenses", budgetTotalExpenses, actualTotalExpenses, totalExpenseVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Profit", budgetMonthProfit, cashOnlyProfit, monthIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n%n", "Paying Students", payingStudentsBudget, payingStudentsActual, payingStudentsVariance);
        System.out.println("(10) Finished computing Budget Sheet Entries");
    }

    /******************************************************************************************
     * Update Budget Excel Workbook
     ******************************************************************************************/
    public void updateBudgetWorkbook(XSSFWorkbook budgetWorkbook, int targetMonthColumnIndex)
    {
        System.out.println("(11) Start updating budget XSSFsheet");
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
                        row.getCell(targetMonthColumnIndex).setCellValue(actualCashCombinedDirectPublicSupport);
                        row.getCell(13).setCellValue(directPublicSupportVariance);
                        break;
                    case "Cash Combined Tuition Fees":
                        row.getCell(targetMonthColumnIndex).setCellValue(budgetCashCombinedTuitionFees);
                        row.getCell(13).setCellValue(tuitionFeeVariance);
                        break;
                    case "Cash Misc Income":
                        row.getCell(targetMonthColumnIndex).setCellValue(actualCashMiscIncome);
                        break;
                    case "Cash Only Income":
                        row.getCell(targetMonthColumnIndex).setCellValue(budgetCashOnlyIncome);
                        row.getCell(13).setCellValue(totalIncomeVariance);
                        break;
                    case "Cash Combined Salaries":
                        row.getCell(targetMonthColumnIndex).setCellValue(actualCashCombinedSalaries);
                        row.getCell(13).setCellValue(monthSalaryVariance);
                        break;
                    case "Cash Combined Contract Services":
                        row.getCell(targetMonthColumnIndex).setCellValue(actualCashCombinedContractServices);
                        row.getCell(13).setCellValue(contractServiceVariance);
                        break;
                    case "Cash Combined Facilities and Equipment":
                        row.getCell(targetMonthColumnIndex).setCellValue(budgetCashCombinedFacilitiesEquiment);
                        row.getCell(13).setCellValue(facilitiesVariance);
                        break;
                    case "Cash Combined Operations":
                        row.getCell(targetMonthColumnIndex).setCellValue(actualCashCombinedOperations);
                        row.getCell(13).setCellValue(operationsVariance);
                        break;
                        case "Cash Misc Expenses":
                        row.getCell(targetMonthColumnIndex).setCellValue(actualCashMiscExpenses);
                        //row.getCell(13).setCellValue(actualCashMiscExpenses);
                        break;
                    case "Cash Only Expenses":
                        row.getCell(targetMonthColumnIndex).setCellValue(budgetCashOnlyExpenses);
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
        System.out.println("(12) Finished updating budget XSSFsheet");
    }
}



