package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200823
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
    int targetMonthColumnIndex;
    private int pandlContractServices;
    private int budgetTotalExpenses;
    private int totalIncomeVariance;
    private int pandlSalaries;
    private int budgetSalaries;
    private int budgetContractServices;
    private int pandlPayrollServiceFees;
    private int contractServiceVariance;
    private int pandlDepreciation;
    private int pandlOperations;
    private int pandlTravel;
    private int budgetOperations;
    private int operationsVariance;
    private int totalExpenseVariance;
    private int payingStudentsVariance;
    private int budgetTotalIncome;
    private int budgetProfit;
    private int cashOnlyProfit;
    private int pandlContributedServices;
    private int pandlGiftsInKindGoods;
    private int cashMiscIncome;
    private int actualCashMiscExpenses;
    private int pandlProgramIncome;
    private int grantsAndGifts;
    private int budgetMiscIncome;
    private int pandlMiscExpenses;
    private int pandlOtherIncome;
    private int budgetInvestments;
    private int budgetCashOnlyIncome;
    private int budgetCashTuitionFees;
    private int budgetMiscExpenses;
    private int budgetRent;
    private int cashOnlyExpenses;
    private int pandlTotalGrantScholarship;
    private int pandlOtherExpenses;
    private int pandlLeagueScholarship;
    private int pandlGrantsAndGifts;
    private int pandlTotalProgramIncome;
    private int pandlRent;
    private int pandlMiscIncome;
    private int pandlBreakRoomSupplies;
    private int pandlPenalties;
    private int cashGrantsGifts;
    private int cashTuition;
    private int pandlInvestments;
    private int cashTotalIncome;
    private int cashSalaries;
    private int cashContractServices;
    private int cashRent;
    private int pandlScholarships;
    private int cashOperations;
    private int cashMiscExpenses;
    private int cashTotalExpenses;
    private int cashProfit;
    private int grantsGiftsVariance;
    private int tuitionVariance;
    private int budgetGrantsGifts;
    private int budgetTuition;
    private int miscIncomeVariance;
    private int salaryVariance;
    private int rentVariance;
    private int miscExpenseVariance;
    private int profitVariance;
    private int cashPayingStudents;
    private int budgetPayingStudents;

    /******************************************************************************************
     * Extract Needed Budget Items from P & L Map
     ******************************************************************************************/
    public void extractBudgetItemsFromPandLMap(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetHashMap)
    {
        System.out.println("(5) Extracting Items for Budget From PandL Map");
        for (String key : budgetHashMap.keySet())
        {
            String switchKey = key.trim();
            if (pandLmap.get(switchKey) != null)//Check for missing P&L items
            switch (switchKey)
            {
                case "Grants and Gifts":
                    pandlGrantsAndGifts = pandLmap.get("Total 43400 Direct Public Support");
                    pandlContributedServices = pandLmap.get("43460 Contributed Services");//Non cash item...must be subtracted
                    break;
                case "Tuition":
                    pandlProgramIncome = pandLmap.get("Total 47200 Program Income");
                    pandlLeagueScholarship = pandLmap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
                    break;
                case "Misc Income":
                    pandlMiscIncome = pandLmap.get("Total 45000 Investments");
                    break;
                case "Salaries":
                    pandlSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    pandlContributedServices = pandLmap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
                    break;
                case "Contract Services":
                    pandlContractServices = pandLmap.get("Total 62100 Contract Services");
                    pandlPayrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    break;
                case "Rent":
                    pandlRent = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");//Non cash item...must be subtracted
                    break;
                case "Operations":
                    pandlOperations = pandLmap.get("Total 65000 Operations");
                    pandlBreakRoomSupplies = pandLmap.get("65055 Breakroom Supplies");
                    pandlOtherExpenses = pandLmap.get("65100 Other Types of Expenses");
                    pandlTravel = pandLmap.get("Total 68300 Travel and Meetings");
                    pandlScholarships = pandLmap.get("65090 Scholarships");
                    break;
                case "Misc Expenses":
                    pandlPenalties = pandLmap.get("90100 Penalties");
                    break;
                default:
                    break;
            }
        }
        System.out.println("(6) Finished Extracting Items for Budget From PandL Map");
    }
    /******************************************************************************************
     * Extract Needed Budget Items from Budget Map
     ******************************************************************************************/
    public void extractBudgetItemsFromBudgetMap(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, int targetMonthColumnIndex)
    {
        System.out.println("(7) Extracting Items for Budget  From Budget Map");
        this.targetMonthColumnIndex = targetMonthColumnIndex;
        for (String key : budgetMap.keySet())
        {
            String switchKey = key.trim();
            if (switchKey != null)
            {
                switch (switchKey)
                {
                    case "Grants and Gifts":
                        budgetGrantsGifts = budgetMap.get("Grants and Gifts");
                        break;
                    case "Tuition":
                        budgetTuition = budgetMap.get("Tuition");
                        break;
                    case "Misc Income":
                        budgetMiscIncome = budgetMap.get("Misc Income");
                        break;
                    case "Total Income":
                        budgetTotalIncome = budgetMap.get("Total Income");
                        break;
                    case "Salaries":
                        budgetSalaries = budgetMap.get("Salaries");
                        break;
                    case "Contract Services":
                        budgetContractServices = budgetMap.get("Contract Services");
                        break;
                    case "Rent":
                        budgetRent = budgetMap.get("Rent");
                        break;
                    case "Operations":
                        budgetOperations = budgetMap.get("Operations");
                        break;
                    case "Misc Expenses":
                        budgetMiscExpenses = budgetMap.get("Misc Expenses");
                        break;
                    case "Total Expenses":
                        budgetTotalExpenses = budgetMap.get("Total Expenses");
                        break;
                    case "Profit":
                        budgetProfit = budgetMap.get("Profit");
                        break;
                    case "Profit Variance":
                        if (budgetMap.get("Profit Variance") != null)
                        {
                            profitVariance = budgetMap.get("Profit Variance");
                        }
                        break;
                    case "Paying Students (Actual)":
                        cashPayingStudents = budgetMap.get("Paying Students (Actual)");
                        break;
                    case "Paying Students (Budget)":
                        budgetPayingStudents = budgetMap.get("Paying Students (Budget)");
                        break;
                    default:
                }
            }

        }
        System.out.println("(8) Finished Extracting Items for Budget From Budget Map");
    }
    /******************************************************************************************
     * Compute budget sheet entries
     ******************************************************************************************/
    public void computeCombinedCashBudgetSheetEntries()
    {
        System.out.println("(9) Computing Combined Budget Sheet Entries");
        cashGrantsGifts = pandlGrantsAndGifts - pandlContributedServices;
        grantsGiftsVariance = cashGrantsGifts - budgetGrantsGifts;
        cashTuition = pandlProgramIncome - pandlLeagueScholarship;
        tuitionVariance = cashTuition - budgetTuition;
        cashMiscIncome = pandlMiscIncome;
        miscIncomeVariance = cashMiscIncome - budgetMiscIncome;
        cashTotalIncome = cashGrantsGifts + cashTuition + cashMiscIncome;
        totalIncomeVariance = cashTotalIncome - budgetTotalIncome;
        cashSalaries = pandlSalaries + pandlPayrollServiceFees - pandlContributedServices;
        salaryVariance = cashSalaries - budgetSalaries;
        cashContractServices = pandlContractServices;
        contractServiceVariance = cashContractServices - budgetContractServices;
        cashRent = pandlRent - pandlDepreciation;
        rentVariance = cashRent - budgetRent;
        cashOperations = pandlOperations + pandlBreakRoomSupplies + pandlOtherExpenses + pandlTravel - pandlScholarships;
        operationsVariance = cashOperations - budgetOperations;
        cashMiscExpenses = pandlPenalties;
        miscExpenseVariance = cashMiscExpenses - budgetMiscExpenses;
        cashTotalExpenses = cashSalaries + cashContractServices + cashRent + cashOperations + cashMiscExpenses;
        totalExpenseVariance = cashTotalExpenses - budgetTotalExpenses;
        cashProfit = cashTotalIncome - cashTotalExpenses;
        profitVariance = cashProfit - budgetProfit;
        payingStudentsVariance = cashPayingStudents - budgetPayingStudents;
        /***************************************************************************************************************
         * Print Budget Proof Figures
         ***************************************************************************************************************/
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "BUDGET AMOUNT", "CASH AMOUNT", "VARIANCE for Month " + targetMonthColumnIndex);
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "---------------------");
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Grants and Gifts", budgetGrantsGifts, cashGrantsGifts, grantsGiftsVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition", budgetTuition, cashTuition, tuitionVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Misc Income", budgetMiscIncome, cashMiscIncome, miscIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", budgetTotalIncome, cashTotalIncome, totalIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", budgetSalaries, cashSalaries, salaryVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Contract Services", budgetContractServices, cashContractServices, contractServiceVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Rent",  budgetRent, cashRent, rentVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", budgetOperations, cashOperations, operationsVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Misc Expenses", budgetMiscExpenses, cashMiscExpenses, miscExpenseVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Expenses", budgetTotalExpenses, cashTotalExpenses, totalExpenseVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Profit", budgetProfit, cashProfit, profitVariance);
        System.out.printf("%-40s %-20s %-20s %,-20d %n", "Profit Variance", "-", "-", profitVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n%n", "Paying Students", budgetPayingStudents, cashPayingStudents, payingStudentsVariance);
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
                    case "Grants and Gifts":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashGrantsGifts);
                        row.getCell(13).setCellValue(grantsGiftsVariance);
                        break;
                    case "Tuition":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashTuition);
                        row.getCell(13).setCellValue(tuitionVariance);
                        break;
                    case "Misc Income":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashMiscIncome);
                        row.getCell(13).setCellValue(miscIncomeVariance);
                        break;
                    case "Total Income":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashTotalIncome);
                        row.getCell(13).setCellValue(totalIncomeVariance);
                        break;
                    case "Salaries":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashSalaries);
                        row.getCell(13).setCellValue(salaryVariance);
                        break;
                    case "Contract Services":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashContractServices);
                        row.getCell(13).setCellValue(contractServiceVariance);
                        break;
                    case "Rent":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashRent);
                        row.getCell(13).setCellValue(rentVariance);
                        break;
                    case "Operations":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashOperations);
                        row.getCell(13).setCellValue(operationsVariance);
                        break;
                        case "Misc Expenses":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashMiscExpenses);
                        row.getCell(13).setCellValue(miscExpenseVariance);
                        break;
                    case "Total Expenses":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashTotalExpenses);
                        row.getCell(13).setCellValue(totalExpenseVariance);
                        break;
                    case "Profit":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashProfit);
                        row.getCell(13).setCellValue(profitVariance);
                        break;
                     case "Profit Variance":
                        row.getCell(targetMonthColumnIndex).setCellValue(profitVariance);
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



