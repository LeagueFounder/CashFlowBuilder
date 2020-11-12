package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201111
// * copyright 2020 Vic Wintriss
//******************************************************************************************
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.LocalDateTime;
import java.util.HashMap;
public class CashItemAggregator
{
    private int targetMonth;
    private int pandLContractServices = 0;
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
    private int pandLTotalIncome;
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
    public CashItemAggregator(HashMap<String, Integer> budgetMap, HashMap<String, Integer> pandLMap, int targetMonth)
    {
        this.budgetMap = budgetMap;
        this.pandLMap = pandLMap;
        this.targetMonth = targetMonth;
        System.out.println("(5) Computing Combined Budget Sheet Entries");
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "---------------------");
        computeGrantsAndGifts();
        computeTuition();
        computeSalaries();
        computeTotalIncome();
        computeContractServices();
        computeContractServices();
        computeRent();
        computeOperatons();
        computeTotalExpenses();
        computeProfit();
        computeStudents();
        reconcile();
    }
    public void computeGrantsAndGifts()
    {
        int budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        int pandLDirectPublicSupport = pandLMap.get("Total 43400 Direct Public Support");
        int pandLGiftsInKindGoods = pandLMap.get("43440 Gifts in Kind - Goods");//Non cash item...must be subtracted
        int pandLContributedServices = pandLMap.get("43460 Contributed Services");//Non cash item...must be subtracted
        int pandLInterest = pandLMap.get("Total 45000 Investments");
        pandLGrantsGifts = pandLDirectPublicSupport - pandLContributedServices - pandLGiftsInKindGoods + pandLInterest;
        grantsGiftsVariance = pandLGrantsGifts - budgetGrantsGifts;
        printConsoleSummary("Grants and Gifts", budgetGrantsGifts, pandLGrantsGifts, grantsGiftsVariance);
    }
    public void computeTuition()
    {
        int pandLProgramIncome = pandLMap.get("Total 47200 Program Income");
        int pandLLeagueScholarship = pandLMap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
        int budgetTuition = budgetMap.get("Tuition");
        pandLTuition = pandLProgramIncome - pandLLeagueScholarship;
        tuitionVariance = pandLTuition - budgetTuition;
        printConsoleSummary("Tuition", budgetTuition, pandLTuition, tuitionVariance);
    }
    public void computeSalaries()
    {
        pandLSalaries = pandLMap.get("Total 62000 Salaries & Related Expenses");
        int pandLPayrollServiceFees = pandLMap.get("62145 Payroll Service Fees");
        int budgetSalaries = budgetMap.get("Salaries");
        int pandLContributedServices = pandLMap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
        pandLSalaries = pandLSalaries + pandLPayrollServiceFees - pandLContributedServices;
        salaryVariance = pandLSalaries - budgetSalaries;
        printConsoleSummary("Salaries", budgetSalaries, pandLSalaries, salaryVariance);
    }
    public void computeTotalIncome()
    {
        pandLTotalIncome = pandLGrantsGifts + pandLTuition;
        int budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        int budgetTuition = budgetMap.get("Tuition");
        budgetTotalIncome = budgetGrantsGifts + budgetTuition;
        incomeTotalVariance = pandLTotalIncome - budgetTotalIncome;
        printConsoleSummary("Total Income", budgetTotalIncome, pandLTotalIncome, incomeTotalVariance);
    }
    public void computeContractServices()
    {
        pandLContractServices = pandLMap.get("Total 62100 Contract Services");
        budgetContractServices = budgetMap.get("Contract Services");
        contractServiceVariance = pandLContractServices - budgetContractServices;
        printConsoleSummary("Contract Services", budgetContractServices, pandLContractServices, contractServiceVariance);
    }
    public void computeRent()
    {
        pandLDepreciation = pandLMap.get("62810 Depr and Amort - Allowable");//Non cash item...must be subtracted
        budgetRent = budgetMap.get("Rent");
        pandLRent = pandLMap.get("Total 62800 Facilities and Equipment");
        pandLRent = pandLRent;// - pandLDepreciation;
        rentVariance = pandLRent - budgetRent;
        printConsoleSummary("Rent", budgetRent, pandLRent, rentVariance);
    }
    public void computeOperatons()
    {
        budgetOperations = budgetMap.get("Operations");
        pandLTravel = pandLMap.get("Total 68300 Travel and Meetings");
        pandLOtherExpenses = pandLMap.get("Total 65100 Other Types of Expenses");
        pandLBreakRoomSupplies = pandLMap.get("65055 Breakroom Supplies");
        pandLOperations = pandLMap.get("Total 65000 Operations");
        pandLOperations = pandLOperations + pandLBreakRoomSupplies + pandLOtherExpenses + pandLTravel;
        operationsVariance = pandLOperations - budgetOperations;
        printConsoleSummary("Operations", budgetOperations, pandLOperations, operationsVariance);
    }
    public void computeTotalExpenses()
    {
        budgetTotalExpenses = budgetMap.get("Total Expenses");
        pandLTotalExpenses = pandLSalaries + pandLContractServices + pandLRent + pandLOperations;
        expenseTotalVariance = pandLTotalExpenses - budgetTotalExpenses;
        printConsoleSummary("Total Expenses", budgetTotalExpenses, pandLTotalExpenses, expenseTotalVariance);
    }
    public void computeProfit()
    {
        budgetProfit = budgetMap.get("Profit");
        pandLProfit = pandLTotalIncome - pandLTotalExpenses;
        profitVariance = pandLProfit - budgetProfit;
        printConsoleSummary("Profit", budgetProfit, pandLProfit, profitVariance);
    }
    public void computeStudents()
    {
        budgetPayingStudents = budgetMap.get("Paying Students");
        actualPayingStudents = pandLTuition / 240;//Derived...including workshops, slams, etc and partial paying students
        payingStudentsVariance = actualPayingStudents - budgetPayingStudents;
        printConsoleSummary("Paying Students", budgetPayingStudents, actualPayingStudents, payingStudentsVariance);
    }
    public void reconcile()
    {
        pandLBottomLineExpense = pandLMap.get("Total Expenses");
        pandLBottomLineIncome = pandLMap.get("Total Income");
        pandLBottomLineProfit = pandLMap.get("Net Income");//Take out in-kind donations!
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
    public void printConsoleSummary(String account, int budgetAmount, int actualAmount, int variance)
    {
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", account, budgetAmount, actualAmount, variance);
    }
    //******************************************************************************************
    //* Update Budget Excel Workbook
    //******************************************************************************************
    public void updateBudgetWorkbook(XSSFWorkbook budgetWorkbook, int targetMonth)
    {
        System.out.println("(7) Start updating budget XSSFsheet");
        LocalDateTime now = LocalDateTime.now();
        XSSFSheet budgetSheet = budgetWorkbook.getSheetAt(0);
        for (Row row : budgetSheet)
        {
            if (row.getRowNum() == 0 || row.getRowNum() ==1)
            {
                row.createCell(13, XSSFCell.CELL_TYPE_STRING);//For month variance numbers
                try
                {
                    budgetSheet.getRow(0).getCell(0).setCellValue("Updated: " + now);
                    budgetSheet.getRow(0).getCell(13).setCellValue("Month " + targetMonth);
                    budgetSheet.getRow(1).getCell(13).setCellValue("VARIANCE");
                    budgetSheet.getRow(1).getCell(targetMonth).setCellValue(">>ACTUAL<");
                }
                catch (Exception e)
                {

                }
            }
            else
            {
                row.createCell(13, XSSFCell.CELL_TYPE_NUMERIC);//For month variance numbers
            }

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
        System.out.println("(8) Finished updating budget XSSFsheet");
    }
}
