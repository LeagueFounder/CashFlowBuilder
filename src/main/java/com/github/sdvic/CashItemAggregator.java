package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 201116
// * copyright 2020 Vic Wintriss
//******************************************************************************************
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
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
    private double pandLDonations;
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
    private double pandLGrantScholarships;
    private double miscIncomeVariance;
    private double pandLMiscIncome;
    private double budgetMiscExpense;
    private double pandLMiscExpense;
    private double miscExpenseVariance;
    private double pandLBusinessExpenses;
    private double pandLOtherIncome;
    private double pandLInvestments;
    private double budgetMiscIncome;
    private double updateBudgetProfit;
    public CashItemAggregator(HashMap<String, Integer> budgetMap, HashMap<String, Double> pandLMap, int targetMonth)
    {
        this.budgetMap = budgetMap;
        this.pandLMap = pandLMap;
        this.targetMonth = targetMonth;
        System.out.println("(5) Computing Combined Budget Sheet Entries");
        printConsoleSummary("ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        printConsoleSummary( "------------------------------------", "-------------", "-------------", "---------------------");
        System.out.println();
    }
    public void computeDonations()
    {
        double budgetDonations = budgetMap.get("Donations");
        double pandLDirectPublicSupport = pandLMap.get("Total 43400 Direct Public Support");
        pandLContributedServices = pandLMap.get("43460 Contributed Services");//Non cash item...must be subtracted
        pandLGrantScholarships = pandLMap.get("Total 47204 Grant Scholarship");
        pandLDonations = pandLDirectPublicSupport - pandLContributedServices + pandLGrantScholarships;
        grantsGiftsVariance = pandLDonations - budgetDonations;
        printConsoleSummary("Donations", budgetDonations, pandLDonations, grantsGiftsVariance);
    }
    public void computeTuition()
    {
        double pandLProgramIncome = pandLMap.get("Total 47200 Program Income");
        pandLLeagueScholarship = pandLMap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
        double budgetTuition = budgetMap.get("Tuition");
        pandLTuition = pandLProgramIncome - pandLLeagueScholarship;
        tuitionVariance = pandLTuition - budgetTuition;
        printConsoleSummary("Tuition", budgetTuition, pandLTuition, tuitionVariance);
    }
    public void computeMiscIncome()
    {
        double budgetMiscIncome = budgetMap.get("Misc Income");
        pandLInvestments = pandLMap.get("Total 45000 Investments");
        pandLOtherIncome = pandLMap.get("Total 46400 Other Types of Income");
        pandLMiscIncome = pandLOtherIncome + pandLInvestments;
        miscIncomeVariance = pandLMiscIncome - budgetMiscIncome;
        printConsoleSummary("MiscIncome", budgetMiscIncome, pandLMiscIncome, miscIncomeVariance);
    }
    public void computeTotalIncome()
    {
        double budgetGrantsGifts = budgetMap.get("Donations");
        double budgetTuition = budgetMap.get("Tuition");
        double pandLInvestments = pandLMap.get("Total 45000 Investments");
        double pandLOtherIncome = pandLMap.get("Total 46400 Other Types of Income");
        double pandLMiscIncome = pandLInvestments + pandLOtherIncome;
        pandLTotalIncome = pandLDonations + pandLTuition + pandLMiscIncome - pandLGrantScholarships;//grants in both Tuition and Donations
        budgetTotalIncome = budgetGrantsGifts + budgetTuition + budgetMiscIncome;
        incomeTotalVariance = pandLTotalIncome - budgetTotalIncome;
        printConsoleSummary("Total Income", budgetTotalIncome, pandLTotalIncome, incomeTotalVariance);
    }
    public void computeSalaries()
    {
        pandLSalaries = pandLMap.get("Total 62000 Salaries & Related Expenses");
        double pandLPayrollServiceFees = pandLMap.get("62145 Payroll Service Fees");
        double budgetSalaries = budgetMap.get("Salaries");
        pandLSalaries = pandLSalaries + pandLPayrollServiceFees - pandLContributedServices;
        salaryVariance = pandLSalaries - budgetSalaries;
        printConsoleSummary("Salaries", budgetSalaries, pandLSalaries, salaryVariance);
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
        budgetRent = budgetMap.get("Rent");
        pandLRent = pandLMap.get("Total 62800 Facilities and Equipment");
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
        pandLBusinessExpenses = pandLMap.get("Total 60900 Business Expenses");
        pandLOperations = pandLOperations + pandLBreakRoomSupplies + pandLOtherExpenses + pandLTravel + pandLBusinessExpenses;
        operationsVariance = pandLOperations - budgetOperations;
        printConsoleSummary("Operations", budgetOperations, pandLOperations, operationsVariance);
    }
    public void computeMiscExpense()
    {
        pandLMiscExpense = pandLMap.get("Total Other Expenses");
        budgetMiscExpense = budgetMap.get("Misc Expense");
        miscExpenseVariance = budgetMiscExpense + pandLMiscExpense;
        printConsoleSummary("Misc Expense", budgetMiscExpense, pandLMiscExpense, miscExpenseVariance);

    }
    public void computeTotalExpenses()
    {
        budgetTotalExpenses = budgetMap.get("Total Expenses");
        pandLTotalExpenses = pandLSalaries + pandLContractServices + pandLRent + pandLOperations + pandLMiscExpense;
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
        actualPayingStudents = (int)(pandLTuition / 240);//Derived...including workshops, slams, etc and partial paying students
        payingStudentsVariance = actualPayingStudents - budgetPayingStudents;
        printConsoleSummary("Paying Students", budgetPayingStudents, actualPayingStudents, payingStudentsVariance);
    }
    public void reconcile()
    {
        double pandLExpense = pandLMap.get("Total Expenses");
        pandLIncome = pandLMap.get("Total Income");
        double pandLNetIncome = pandLMap.get("Net Income");
        double pandlIncomeVariance = pandLIncome - pandLIncome;
        pandLTotalExpenses = pandLSalaries + pandLContractServices + pandLRent + pandLOperations;
        double pandlExpenseVariance = pandLBottomLineExpense - pandLTotalExpenses;
        double pandlProfitVariance = pandLProfit - pandLNetIncome;
        printConsoleSummary("", "", "P&L RECONCILIATION", "");
        printConsoleSummary("ACCOUNT", "BUDGET", "P&L", "VARIANCE", "-LSC");
        printConsoleSummary(  "------------------------------------", "------------", "------------", "----------", "---------------");
        System.out.println();
        printConsoleSummary("Income", pandLTotalIncome, pandLIncome, pandlIncomeVariance, pandLLeagueScholarship);
        printConsoleSummary("Expenses", pandLTotalExpenses, pandLBottomLineExpense, pandlExpenseVariance);
        printConsoleSummary("Profit", pandLProfit, pandLNetIncome, pandlProfitVariance);
        System.out.println("(6) Finished computing Budget Sheet Entries");
    }
    public void printConsoleSummary(String title1, double title2, double title3, double title4)
    {
        System.out.printf("%-40s %,-20.0f %,-20.0f %,-20.0f %n", title1, title2, title3, title4);
    }
    public void printConsoleSummary(String title1, double title2, double title3, double title4, double title5)
    {
        System.out.printf("%-40s %,-20.0f %,-20.0f %,-20.0f -20.0f %n", title1, title2, title3, title4, title5);
    }
    public void printConsoleSummary(String title1, String title2, String title3, String title4)
    {
        System.out.printf("%n %-40s %-20s %-20s %-20s", title1, title2, title3, title4);
    }
    public void printConsoleSummary(String title1, String title2, String title3, String title4, String title5)
    {
        System.out.printf("%n %-40s %-20s %-20s %-20s %-40s", title1, title2, title3, title4, title5);
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
                    case "Donations":
                        row.getCell((int) targetMonth).setCellValue(pandLDonations);
                        row.getCell(13).setCellValue((int)grantsGiftsVariance);
                        break;
                    case "Tuition":
                        row.getCell((int) targetMonth).setCellValue(pandLTuition);
                        row.getCell(13).setCellValue(tuitionVariance);
                        break;
                    case "Misc Income":
                        row.getCell((int) targetMonth).setCellValue((int)pandLMiscIncome);
                        row.getCell(13).setCellValue(miscIncomeVariance);
                        break;
                    case "Total Income":
                        row.getCell((int) targetMonth).setCellValue(pandLTotalIncome);
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
                    case "Misc Expense":
                        row.getCell((int) targetMonth).setCellValue(pandLMiscExpense);
                        row.getCell(13).setCellValue((int)miscExpenseVariance);
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
