package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201030
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
    private int pandLContractServices;
    private int pandLSalaries;
    private int contractServiceVariance;
    private int pandLOperations;
    private int operationsVariance;
    private int payingStudentsVariance;
    private int pandLRent;
    private int grantsGiftsVariance;
    private int tuitionVariance;
    private int budgetGrantsGifts;
    private int budgetTuition;
    private int salaryVariance;
    private int rentVariance;
    private int profitVariance;
    private int budgetPayingStudents;
    private int expenseTotalVariance;
    private int pandLTotalIncome;
    private int incomeTotalVariance;
    private int pandLProfit;
    private int pandLIncome;
    private int pandLTotalExpenses;
    private int pandLAccumulatedProfit;
    private int pandLGrantsGifts;
    private int pandLTuition;
    private int actualPayingStudents;
    private int pandLContributedServices;
    private int pandLDirectPublicSupport;
    private int pandLGiftsInKindGoods;
    private int budgetContractServices;
    private int pandLProgramIncome;
    private int pandLLeagueScholarship;
    private int budgetTotalIncome;
    private int pandLPayrollServiceFees;
    private int budgetSalaries;
    private int budgetRent;
    private int pandLBreakRoomSupplies;
    private int pandLOtherExpenses;
    private int pandLTravel;
    private int budgetOperations;
    private int pandLDepreciation;
    private int budgetTotalExpenses;
    private int budgetProfit;
    private int pandLBottomLineProfit;
    private int pandLBottomLineExpense;
    private int pandLBottomLineIncome;
    private HashMap<String, Integer> budgetMap;
    private HashMap<String, Integer> pandLMap;
    public CashItemAggregator(HashMap<String, Integer> budgetMap, HashMap<String, Integer> pandLMap)
    {
        this.budgetMap = budgetMap;
        this.pandLMap = pandLMap;
        computeGrantsAndGifts();
    }
    public void extractNumbersFromExcelSheets()
    {
        try
        {
            budgetGrantsGifts = budgetMap.get("Grants and Gifts");
            pandLContributedServices = pandLMap.get("43460 Contributed Services");//Non cash item...must be subtracted
            pandLSalaries = pandLMap.get("Total 62000 Salaries & Related Expenses");
            pandLContributedServices = pandLMap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
            pandLRent = pandLMap.get("Total 62800 Facilities and Equipment");
            pandLOperations = pandLMap.get("Total 65000 Operations");
            pandLTotalExpenses = pandLMap.get("Total Expenses");
            budgetMap.get("Profit");
            budgetPayingStudents = budgetMap.get("Paying Students");
            budgetMap.get("Contract Services");
            pandLContractServices = pandLMap.get("Total 62100 Contract Services");
            pandLOperations = pandLMap.get("Total 65000 Operations");
            pandLBreakRoomSupplies = pandLMap.get("65055 Breakroom Supplies");
            pandLOtherExpenses = pandLMap.get("Total 65100 Other Types of Expenses");
            pandLTravel = pandLMap.get("Total 68300 Travel and Meetings");
            budgetOperations = budgetMap.get("Operations");
            pandLDirectPublicSupport = pandLMap.get("Total 43400 Direct Public Support");
            pandLGiftsInKindGoods = pandLMap.get("43440 Gifts in Kind - Goods");//Non cash item...must be subtracted
            budgetContractServices = budgetMap.get("Contract Services");
            pandLContractServices = pandLMap.get("Total 62100 Contract Services");
            pandLProgramIncome = pandLMap.get("Total 47200 Program Income");
            pandLLeagueScholarship = pandLMap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
            budgetTuition = budgetMap.get("Tuition");
            pandLSalaries = pandLMap.get("Total 62000 Salaries & Related Expenses");
            pandLContributedServices = pandLMap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
            pandLPayrollServiceFees = pandLMap.get("62145 Payroll Service Fees");
            budgetSalaries = budgetMap.get("Salaries");
            pandLRent = pandLMap.get("Total 62800 Facilities and Equipment");
            budgetRent = budgetMap.get("Rent");
            pandLDepreciation = pandLMap.get("62810 Depr and Amort - Allowable");//Non cash item...must be subtracted
            pandLContractServices = pandLMap.get("Total 62100 Contract Services");
            budgetContractServices = budgetMap.get("Contract Services");
            budgetTotalExpenses = budgetMap.get("Total Expenses");
            pandLTotalExpenses = pandLMap.get("Total Expenses");
            budgetProfit = budgetMap.get("Profit");
            budgetPayingStudents = budgetMap.get("Paying Students");
            pandLBottomLineProfit = pandLMap.get("Net Income");//Take out in-kind donations!
            pandLBottomLineExpense = pandLMap.get("Total Expenses");
            pandLBottomLineIncome = pandLMap.get("Total Income");
        }
        catch (Exception e)
        {
            System.out.println("Error reading Excel sheets in extractNumbersFromExcelSheets()  ");
        }
    }
    public void computeGrantsAndGifts()
    {
        budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        pandLContributedServices = pandLMap.get("43460 Contributed Services");//Non cash item...must be subtracted
        pandLDirectPublicSupport = pandLMap.get("Total 43400 Direct Public Support");
        pandLGiftsInKindGoods = pandLMap.get("43440 Gifts in Kind - Goods");//Non cash item...must be subtracted
        pandLGrantsGifts = pandLDirectPublicSupport - pandLContributedServices - pandLGiftsInKindGoods;
        grantsGiftsVariance = pandLGrantsGifts - budgetGrantsGifts;
        printConsoleSummary("Grants and Gifts", budgetGrantsGifts, pandLGrantsGifts, grantsGiftsVariance);
    }
    public void computeCombinedCashBudgetSheetEntries(XSSFWorkbook budgetWorkBook, int targetMonth)
    {
        this.budgetMap = budgetMap;
        System.out.println("(5) Computing Combined Budget Sheet Entries");
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "---------------------");
        //*************************************************************************************************************
        //* TUITION
        //*************************************************************************************************************
        pandLTuition = pandLProgramIncome - pandLLeagueScholarship;
        tuitionVariance = pandLTuition - budgetTuition;
        printConsoleSummary("Tuition", budgetTuition, pandLTuition, tuitionVariance);
        //*************************************************************************************************************
        //* TOTAL INCOME
        //**************************************************************************************************************
        pandLTotalIncome = pandLGrantsGifts + pandLTuition;
        budgetTotalIncome = budgetGrantsGifts + budgetTuition;
        incomeTotalVariance = pandLTotalIncome - budgetTotalIncome;
        printConsoleSummary("Total Income", budgetTotalIncome, pandLTotalIncome, incomeTotalVariance);
        //*************************************************************************************************************
        //* SALARIES
        //**************************************************************************************************************
        pandLSalaries = pandLSalaries + pandLPayrollServiceFees - pandLContributedServices;
        salaryVariance = pandLSalaries - budgetSalaries;
        printConsoleSummary("Salaries", budgetSalaries, pandLSalaries, salaryVariance);
        //*************************************************************************************************************
        //* CONTRACT SERVICES
        //*************************************************************************************************************
        contractServiceVariance = pandLContractServices - budgetContractServices;
        printConsoleSummary("Contract Services", budgetContractServices, pandLContractServices, contractServiceVariance);
        //*************************************************************************************************************
        //* RENT
        //*************************************************************************************************************
        pandLRent = pandLRent - pandLDepreciation;
        rentVariance = pandLRent - budgetRent;
        printConsoleSummary("Rent", budgetRent, pandLRent, rentVariance);
        //*************************************************************************************************************
        //* OPERATIONS
        //*************************************************************************************************************
        pandLOperations = pandLOperations + pandLBreakRoomSupplies + pandLOtherExpenses + pandLTravel;
        operationsVariance = pandLOperations - budgetOperations;
        printConsoleSummary("Operations", budgetOperations, pandLOperations, operationsVariance);
        //*************************************************************************************************************
        //* TOTAL EXPENSES
        //*************************************************************************************************************
        pandLTotalExpenses = pandLSalaries + pandLContractServices + pandLRent + pandLOperations;
        expenseTotalVariance = pandLTotalExpenses - budgetTotalExpenses;
        printConsoleSummary("Total Expenses", budgetTotalExpenses, pandLTotalExpenses, expenseTotalVariance);
        //*************************************************************************************************************
        //* PROFIT
        //*************************************************************************************************************
        pandLProfit = pandLTotalIncome - pandLTotalExpenses;
        profitVariance = budgetProfit - pandLProfit;
        profitVariance = pandLProfit - budgetProfit;
        printConsoleSummary("Profit", budgetProfit, pandLProfit, profitVariance);
        //*************************************************************************************************************
        //* STUDENTS
        // *************************************************************************************************************
        actualPayingStudents = pandLTuition / 240;//Derived...including workshops, slams, etc and partial paying students
        payingStudentsVariance = actualPayingStudents - budgetPayingStudents;
        printConsoleSummary("Paying Students", budgetPayingStudents, actualPayingStudents, payingStudentsVariance);
        //*************************************************************************************************************
        //* RECONCILE...Profit variance will equal depreciation, which is disregarded in these numbers
        //*************************************************************************************************************
        pandLIncome = pandLGrantsGifts + pandLTuition;
        int pandlIncomeVariance = pandLBottomLineIncome - pandLIncome;
        pandLTotalExpenses = pandLSalaries + pandLContractServices + pandLRent + pandLOperations;
        int pandlExpenseVariance = pandLBottomLineExpense - pandLTotalExpenses;
        int pandlProfitVariance = pandLProfit - pandLBottomLineProfit;
        System.out.printf("%n %76s", "P&L RECONCILIATION");
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "ACCUMULATED", "BOTTOM LINE", "VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "------------", "------------", "----------");
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Income", pandLIncome, pandLBottomLineIncome, pandlIncomeVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Expenses", pandLTotalExpenses, pandLBottomLineExpense, pandlExpenseVariance);
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n%n", "Profit", pandLProfit, pandLBottomLineProfit, pandlProfitVariance);
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
                        row.getCell(targetMonth).setCellValue(pandLGrantsGifts);
                        row.getCell(13).setCellValue(grantsGiftsVariance);
                        break;
                    case "Tuition":
                        row.getCell(targetMonth).setCellValue(pandLTuition);
                        row.getCell(13).setCellValue(tuitionVariance);
                        break;
                    case "Total Income":
                        row.getCell(targetMonth).setCellValue(pandLIncome);
                        row.getCell(13).setCellValue(incomeTotalVariance);
                        break;
                    case "Salaries":
                        row.getCell(targetMonth).setCellValue(pandLSalaries);
                        row.getCell(13).setCellValue(salaryVariance);
                        break;
                    case "Contract Services":
                        row.getCell(targetMonth).setCellValue(pandLContractServices);
                        row.getCell(13).setCellValue(contractServiceVariance);
                        break;
                    case "Rent":
                        row.getCell(targetMonth).setCellValue(pandLRent);
                        row.getCell(13).setCellValue(rentVariance);
                        break;
                    case "Operations":
                        row.getCell(targetMonth).setCellValue(pandLOperations);
                        row.getCell(13).setCellValue(operationsVariance);
                        break;
                    case "Total Expenses":
                        row.getCell(targetMonth).setCellValue(pandLTotalExpenses);
                        row.getCell(13).setCellValue(expenseTotalVariance);
                        break;
                    case "Profit":
                        row.getCell(targetMonth).setCellValue(pandLProfit);
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
    //******************************************************************************************
    //* Print Console Summary
    //******************************************************************************************
    public void printConsoleSummary(String account, int budgetAmount, int actualAmount, int variance)
    {
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", account, budgetAmount, actualAmount, variance);
    }
}
