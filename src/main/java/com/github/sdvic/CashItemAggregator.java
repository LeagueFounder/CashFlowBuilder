package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201009
// * copyright 2020 Vic Wintriss
//******************************************************************************************
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
    private int pandlContractServices;
    private int pandlSalaries;
    private int contractServiceVariance;
    private int pandlOperations;
    private int operationsVariance;
    private int payingStudentsVariance;
    private int pandlRent;
    private int grantsGiftsVariance;
    private int tuitionVariance;
    private int budgetGrantsGifts;
    private int budgetTuition;
    private int salaryVariance;
    private int rentVariance;
    private int profitVariance;
    private int budgetPayingStudents;
    private int expenseTotalVariance;
    private int pandlTotalIncome;
    private int incomeTotalVariance;
    private int pandlProfit;
    private int pandlIncome;
    private int pandlTotalExpenses;
    private int pandlAccumulatedProfit;
    private int pandlGrantsAndGifts, pandlTuition;
    private int actualPayingStudents;
    public void computeCombinedCashBudgetSheetEntries(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, XSSFWorkbook budgetWorkBook, int targetMonth)
    {
        System.out.println("(5) Computing Combined Budget Sheet Entries");
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "---------------------");
        //*************************************************************************************************************
        //* GRANTS AND GIFTS
        //*************************************************************************************************************
        int pandlContributedServices;
        budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        try
        {
            int pandlDirectPublicSupport = pandLmap.get("Total 43400 Direct Public Support");
            pandlContributedServices = pandLmap.get("43460 Contributed Services");//Non cash item...must be subtracted
            int pandlGiftsInKindGoods = pandLmap.get("43440 Gifts in Kind - Goods");//Non cash item...must be subtracted
            pandlGrantsAndGifts = pandlDirectPublicSupport - pandlContributedServices - pandlGiftsInKindGoods;
            grantsGiftsVariance = pandlGrantsAndGifts - budgetGrantsGifts;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Grants and Gifts", budgetGrantsGifts, pandlGrantsAndGifts, grantsGiftsVariance);
        }
        catch(Exception e)
        {
            System.out.println("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ~Line 74, Error processing trying to process GRANTS AND GIFTS, exception => " + e.getMessage());
        }
        //*************************************************************************************************************
        //* TUITION
        //*************************************************************************************************************
        try
        {
            int pandlProgramIncome = pandLmap.get("Total 47200 Program Income");
            int pandlLeagueScholarship = pandLmap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
            budgetTuition = budgetMap.get("Tuition");
            pandlTuition = pandlProgramIncome - pandlLeagueScholarship;
            tuitionVariance = pandlTuition - budgetTuition;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition", budgetTuition, pandlTuition, tuitionVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process TUITION");
        }
        //*************************************************************************************************************
        //* TOTAL INCOME
        //**************************************************************************************************************
        try
        {
            pandlTotalIncome = pandlGrantsAndGifts + pandlTuition;
            int budgetTotalIncome = budgetGrantsGifts + budgetTuition;
            incomeTotalVariance = pandlTotalIncome - budgetTotalIncome;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", budgetTotalIncome, pandlTotalIncome, incomeTotalVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process TOTAL INCOME");
        }
        //*************************************************************************************************************
        //* SALARIES
        //**************************************************************************************************************
        try
        {
            pandlSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
            pandlContributedServices = pandLmap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
            int pandlPayrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
            int budgetSalaries = budgetMap.get("Salaries");
            pandlSalaries = pandlSalaries + pandlPayrollServiceFees - pandlContributedServices;
            salaryVariance = pandlSalaries - budgetSalaries;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", budgetSalaries, pandlSalaries, salaryVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process SALARIES");
        }
        //*************************************************************************************************************
        //* CONTRACT SERVICES
        //*************************************************************************************************************
        try
        {
            pandlContractServices = pandLmap.get("Total 62100 Contract Services");
            int budgetContractServices = budgetMap.get("Contract Services");
            contractServiceVariance = pandlContractServices - budgetContractServices;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Contract Services", budgetContractServices, pandlContractServices, contractServiceVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process CONTRACT SERVICES");
        }
        //*************************************************************************************************************
        //* RENT
        //*************************************************************************************************************
        try
        {
            pandlRent = pandLmap.get("Total 62800 Facilities and Equipment");
            int budgetRent = budgetMap.get("Rent");
            int pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");//Non cash item...must be subtracted
            pandlRent = pandlRent - pandlDepreciation;
            rentVariance = pandlRent - budgetRent;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Rent", budgetRent, pandlRent, rentVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process RENT");
        }
        //*************************************************************************************************************
        //* OPERATIONS
        //*************************************************************************************************************
        try
        {
            pandlOperations = pandLmap.get("Total 65000 Operations");
            int pandlBreakRoomSupplies = pandLmap.get("65055 Breakroom Supplies");
            int pandlOtherExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
            int pandlTravel = pandLmap.get("Total 68300 Travel and Meetings");
            int budgetOperations = budgetMap.get("Operations");
            pandlOperations = pandlOperations + pandlBreakRoomSupplies + pandlOtherExpenses + pandlTravel;
            operationsVariance = pandlOperations - budgetOperations;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", budgetOperations, pandlOperations, operationsVariance);
        }
             catch(Exception e)
        {
            System.out.println("~ Line 218 !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process OPERATIONS, excepetion => " + e.getMessage());
        }
        //*************************************************************************************************************
        //* TOTAL EXPENSES
        //*************************************************************************************************************
        try
        {
            int budgetTotalExpenses = budgetMap.get("Total Expenses");
            pandlTotalExpenses = pandLmap.get("Total Expenses");
            pandlTotalExpenses = pandlSalaries + pandlContractServices + pandlRent + pandlOperations;
            expenseTotalVariance = pandlTotalExpenses - budgetTotalExpenses;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Expenses", budgetTotalExpenses, pandlTotalExpenses, expenseTotalVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process TOTAL EXPENSES");
        }
        //*************************************************************************************************************
        //* PROFIT
        //*************************************************************************************************************
        try
        {
            int budgetProfit = budgetMap.get("Profit");
            pandlProfit = pandlTotalIncome - pandlTotalExpenses;
            profitVariance = budgetProfit - pandlProfit;
            profitVariance = pandlProfit - budgetProfit;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Profit", budgetProfit, pandlProfit, profitVariance);
        }
             catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process PROFIT");
        }
        //*************************************************************************************************************
        //* STUDENTS
        // *************************************************************************************************************
        try
        {
            actualPayingStudents = pandlTuition/240;//Derived...including workshops, slams, etc and partial paying students
            budgetPayingStudents = budgetMap.get("Paying Students");
            payingStudentsVariance = actualPayingStudents - budgetPayingStudents;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Paying Students", budgetPayingStudents, actualPayingStudents, payingStudentsVariance);
        }
             catch(Exception e)
        {
            System.out.println("~ Line 276 !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process STUDENTS, excepetion => " + e.getMessage());
        }
        //*************************************************************************************************************
        //* RECONCILE...Profit vairance will equal depreciation, which is disregarded in these numbers
        //*************************************************************************************************************
        try
        {
            int pandlBottomLineProfit = pandLmap.get("Net Income");//Take out in-kind donations!
            int pandlBottomLineExpense = pandLmap.get("Total Expenses");
            int pandlBottomLineIncome = pandLmap.get("Total Income");
            pandlIncome = pandlGrantsAndGifts + pandlTuition;
            int pandlIncomeVariance = pandlBottomLineIncome - pandlIncome;
            pandlTotalExpenses = pandlSalaries + pandlContractServices + pandlRent + pandlOperations;
            int pandlExpenseVariance = pandlBottomLineExpense - pandlTotalExpenses;
            int pandlProfitVariance = pandlProfit - pandlBottomLineProfit;
            System.out.printf("%n %76s", "P&L RECONCILIATION");
            System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "ACCUMULATED", "BOTTOM LINE", "VARIANCE");
            System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "------------", "------------", "----------");
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Income", pandlIncome, pandlBottomLineIncome, pandlIncomeVariance);
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Expenses", pandlTotalExpenses, pandlBottomLineExpense, pandlExpenseVariance);
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n%n", "Profit", pandlProfit, pandlBottomLineProfit, pandlProfitVariance);
        }
        catch(Exception e)
        {
            System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Error processing trying to process RECONCILE");
        }
        System.out.println("(6) Finished computing Budget Sheet Entries");
    }
    //******************************************************************************************
    //* Update Budget Excel Workbook
    //******************************************************************************************
    public void updateBudgetWorkbook(XSSFWorkbook budgetWorkbook, int targetMonth)
    {
        System.out.println("(7) Start updating budget XSSFsheet");
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
                        row.getCell(targetMonth).setCellValue(pandlGrantsAndGifts);
                        row.getCell(13).setCellValue(grantsGiftsVariance);
                        break;
                    case "Tuition":
                        row.getCell(targetMonth).setCellValue(pandlTuition);
                        row.getCell(13).setCellValue(tuitionVariance);
                        break;
                    case "Total Income":
                        row.getCell(targetMonth).setCellValue(pandlIncome);
                        row.getCell(13).setCellValue(incomeTotalVariance);
                        break;
                    case "Salaries":
                        row.getCell(targetMonth).setCellValue(pandlSalaries);
                        row.getCell(13).setCellValue(salaryVariance);
                        break;
                    case "Contract Services":
                        row.getCell(targetMonth).setCellValue(pandlContractServices);
                        row.getCell(13).setCellValue(contractServiceVariance);
                        break;
                    case "Rent":
                        row.getCell(targetMonth).setCellValue(pandlRent);
                        row.getCell(13).setCellValue(rentVariance);
                        break;
                    case "Operations":
                        row.getCell(targetMonth).setCellValue(pandlOperations);
                        row.getCell(13).setCellValue(operationsVariance);
                        break;
                    case "Total Expenses":
                        row.getCell(targetMonth).setCellValue(pandlTotalExpenses);
                        row.getCell(13).setCellValue(expenseTotalVariance);
                        break;
                    case "Profit":
                        row.getCell(targetMonth).setCellValue(pandlProfit);
                        row.getCell(13).setCellValue(profitVariance);
                        break;
                    case "Profit Variance":
                        row.getCell(targetMonth).setCellValue(profitVariance);
                        break;
                    case "Paying Students":
                        row.getCell(targetMonth).setCellValue(actualPayingStudents);
                        row.getCell(13).setCellValue(payingStudentsVariance);
                        break;
                    default:
                }
            }
        }
        budgetSheet.getRow(0).getCell(0).setCellValue("Updated: " + date);
        budgetSheet.getRow(0).getCell(13).setCellValue("Month " + targetMonth);
        budgetSheet.getRow(1).getCell(13).setCellValue("VARIANCE");
        budgetSheet.getRow(1).getCell(targetMonth).setCellValue(">ACTUAL<");
        System.out.println("(8) Finished updating budget XSSFsheet");
    }
}