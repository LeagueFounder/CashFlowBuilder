package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201112
// * copyright 2020 Vic Wintriss
//******************************************************************************************
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.LocalDateTime;
import java.util.HashMap;
public class CashItemAggregator
{
    private int targetMonth;
    private double pandLContractServices = 0;
    private double pandLSalaries;
    private double contractServiceVariance;
    private double pandLOperations;
    private double operationsVariance;
    private double payingStudentsVariance;
    private double pandLRent;
    private double grantsGiftsVariance;
    private double tuitionVariance;
    private double budgetGrantsGifts;
    private double budgetTuition;
    private double salaryVariance;
    private double rentVariance;
    private double profitVariance;
    private double budgetPayingStudents;
    private double expenseTotalVariance;
    private double incomeTotalVariance;
    private double pandLProfit;
    private double pandLIncome;
    private double pandLTotalExpenses;
    private double pandLAccumulatedProfit;
    private double pandLGrantsGifts;
    private double pandLTuition;
    private double actualPayingStudents;
    private double pandLContributedServices;
    private double pandLDirectPublicSupport;
    private double pandLGiftsInKindGoods;
    private double budgetContractServices;
    private double pandLProgramIncome;
    private double pandLLeagueScholarship;
    private double budgetTotalIncome;
    private double pandLTotalIncome;
    private double pandLPayrollServiceFees;
    private double budgetSalaries;
    private double budgetRent;
    private double pandLBreakRoomSupplies;
    private double pandLOtherExpenses;
    private double pandLTravel;
    private double budgetOperations;
    private double pandLDepreciation;
    private double budgetTotalExpenses;
    private double budgetProfit;
    private double pandLBottomLineProfit;
    private double pandLBottomLineExpense;
    private double pandLBottomLineIncome;
    private HashMap<String, Integer> budgetMap;
    private HashMap<String, Double> pandLMap;
    public CashItemAggregator(HashMap<String, Integer> budgetMap, HashMap<String, Double> pandLMap, int targetMonth)
    {
        this.budgetMap = budgetMap;
        this.pandLMap = pandLMap;
        this.targetMonth = targetMonth;
        System.out.println("(5) Computing Combined Budget Sheet Entries");
        printConsoleSummary("ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        printConsoleSummary( "------------------------------------", "-------------", "-------------", "---------------------");
        System.out.println();
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
        double budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        double pandLDirectPublicSupport = pandLMap.get("Total 43400 Direct Public Support");
        double pandLGiftsInKindGoods = pandLMap.get("43440 Gifts in Kind - Goods");//Non cash item...must be subtracted
        double pandLContributedServices = pandLMap.get("43460 Contributed Services");//Non cash item...must be subtracted
        double pandLInterest = pandLMap.get("Total 45000 Investments");
        pandLGrantsGifts = pandLDirectPublicSupport - pandLContributedServices - pandLGiftsInKindGoods + pandLInterest;
        grantsGiftsVariance = pandLGrantsGifts - budgetGrantsGifts;
        printConsoleSummary("Grants and Gifts", budgetGrantsGifts, pandLGrantsGifts, grantsGiftsVariance);
    }
    public void computeTuition()
    {
        double pandLProgramIncome = pandLMap.get("Total 47200 Program Income");
        double pandLLeagueScholarship = pandLMap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
        double budgetTuition = budgetMap.get("Tuition");
        pandLTuition = pandLProgramIncome - pandLLeagueScholarship;
        tuitionVariance = pandLTuition - budgetTuition;
        printConsoleSummary("Tuition", budgetTuition, pandLTuition, tuitionVariance);
    }
    public void computeSalaries()
    {
        pandLSalaries = pandLMap.get("Total 62000 Salaries & Related Expenses");
        double pandLPayrollServiceFees = pandLMap.get("62145 Payroll Service Fees");
        double budgetSalaries = budgetMap.get("Salaries");
        double pandLContributedServices = pandLMap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
        pandLSalaries = pandLSalaries + pandLPayrollServiceFees - pandLContributedServices;
        salaryVariance = pandLSalaries - budgetSalaries;
        printConsoleSummary("Salaries", budgetSalaries, pandLSalaries, salaryVariance);
    }
    public void computeTotalIncome()
    {
        pandLTotalIncome = pandLGrantsGifts + pandLTuition;
        double budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        double budgetTuition = budgetMap.get("Tuition");
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
        double pandlIncomeVariance = pandLBottomLineIncome - pandLIncome;
        pandLTotalExpenses = pandLSalaries + pandLContractServices + pandLRent + pandLOperations;
        double pandlExpenseVariance = pandLBottomLineExpense - pandLTotalExpenses;
        double pandlProfitVariance = pandLProfit - pandLBottomLineProfit;
        printConsoleSummary("", "", "P&L RECONCILIATION", "");
        printConsoleSummary("ACCOUNT", "ACCUMULATED", "BOTTOM LINE", "VARIANCE");
        printConsoleSummary(  "------------------------------------", "------------", "------------", "----------");
        System.out.println();
        printConsoleSummary("Profit", pandLProfit, pandLBottomLineProfit, pandlProfitVariance);
        printConsoleSummary("Income", pandLIncome, pandLBottomLineIncome, pandlIncomeVariance);
        printConsoleSummary("Expenses", pandLTotalExpenses, pandLBottomLineExpense, pandlExpenseVariance);
        printConsoleSummary("Expenses", pandLTotalExpenses, pandLBottomLineExpense, pandlExpenseVariance);
        System.out.println("(6) Finished computing Budget Sheet Entries");
    }
    public void printConsoleSummary(String account, double budgetAmount, double actualAmount, double variance)
    {
        System.out.printf("%-40s %,-20.0f %,-20.0f %,-20.0f %n", account, budgetAmount, actualAmount, variance);
    }
    public void printConsoleSummary(String title1, String title2, String title3, String title4)
    {
        System.out.printf("%n %-40s %-20s %-20s %-20s", title1, title2, title3, title4);
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
                    budgetSheet.getRow(1).getCell((int) targetMonth).setCellValue(">>ACTUAL<");
                }
                catch (Exception e)
                {

                }
            }
            else
            {
                row.createCell(13, XSSFCell.CELL_TYPE_NUMERIC);//For month variance numbers
                DataFormat format = budgetWorkbook.createDataFormat();
                CellStyle cellstyle = budgetWorkbook.createCellStyle();
                cellstyle.setDataFormat(format.getFormat("#,##0"));
                row.getCell(13).setCellStyle(cellstyle);
            }
            if (row.getCell(0) != null)
            {
                switch (row.getCell(0).getStringCellValue())
                {
                    case "Grants and Gifts":
                        row.getCell((int) targetMonth).setCellValue(pandLGrantsGifts);
                        row.getCell(13).setCellValue((int)grantsGiftsVariance);
                        break;
                    case "Tuition":
                        row.getCell((int) targetMonth).setCellValue(pandLTuition);
                        row.getCell(13).setCellValue(tuitionVariance);
                        break;
                    case "Total Income":
                        row.getCell((int) targetMonth).setCellValue(pandLIncome);
                        row.getCell(13).setCellValue((int)incomeTotalVariance);
                        break;
                    case "Salaries":
                        row.getCell((int) targetMonth).setCellValue(pandLSalaries);
                        row.getCell(13).setCellValue((int)salaryVariance);
                        break;
                    case "Contract Services":
                        row.getCell((int) targetMonth).setCellValue(pandLContractServices);
                        row.getCell(13).setCellValue((int)contractServiceVariance);
                        break;
                    case "Rent":
                        row.getCell((int) targetMonth).setCellValue(pandLRent);
                        row.getCell(13).setCellValue((int)rentVariance);
                        break;
                    case "Operations":
                        row.getCell((int) targetMonth).setCellValue(pandLOperations);
                        row.getCell(13).setCellValue((int)operationsVariance);
                        break;
                    case "Total Expenses":
                        row.getCell((int) targetMonth).setCellValue(pandLTotalExpenses);
                        row.getCell(13).setCellValue((int)expenseTotalVariance);
                        break;
                    case "Profit":
                        row.getCell((int) targetMonth).setCellValue(pandLProfit);
                        row.getCell(13).setCellValue((int)profitVariance);
                        break;
                    case "Profit Variance":
                        row.getCell(targetMonth).setCellValue((int)profitVariance);
                        break;
                    case "Paying Students":
                        row.getCell((int) targetMonth).setCellValue(actualPayingStudents);
                        row.getCell(13).setCellValue((int)payingStudentsVariance);
                        break;
                    default:
                }
            }
        }
        System.out.println("(8) Finished updating budget XSSFsheet");
    }
}
