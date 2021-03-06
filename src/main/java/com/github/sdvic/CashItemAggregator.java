package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 210115
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
    private HashMap<String, Integer> budgetMap;
    private HashMap<String, Double> pandLMap;
    private double pl62100ContractServices;
    private double pl62000Salaries;
    private double contractServiceVariance;
    private double pl65000Operations;
    private double operationsVariance;
    private double payingStudentsVariance;
    private double pl62800Rent;
    private double donationVariance;
    private double tuitionVariance;
    private double budgetGrantsGifts;
    private double budgetMonthlyTuition;
    private double salaryVariance;
    private double rentVariance;
    private double profitVariance;
    private double budgetPayingStudents;
    private double expenseTotalVariance;
    private double incomeTotalVariance;
    private double pandLProfit;
    private double pandLIncome;
    private double pandLTotalExpenses;
    private double pandLDonations;
    private double pandLTuition;
    private double actualPayingStudents;
    private double pandLDirectPublicSupport;
    private double pandLGiftsInKindGoods;
    private double budgetContractServices;
    private double pandLProgramIncome;
    private double pl47203LeagueScholarship;
    private double budgetTotalIncome;
    private double pandLTotalIncome;
    private double budgetSalaries;
    private double budgetRent;
    private double pl65055BreakRoomSupplies;
    private double pl65100OtherExpenses;
    private double pl68300Travel;
    private double budgetOperations;
    private double budgetTotalExpenses;
    private double budgetProfit;
    private double miscIncomeVariance;
    private double pandLMiscIncome;
    private int budgetMiscExpense;
    private double pandLMiscExpense;
    private double miscExpenseVariance;
    private double pl60900BusinessExpenses;
    private double pl46400OtherIncome;
    private double pl45000Investments;
    private int budgetMiscIncome;
    private double pl43400DirectPublicSupport;
    private double pl43460ContributedServices;//Non cash item...musdoublet be subtracted
    private double pl47204GrantScholarships;
    private double pl47200ProgramIncome;
    private double budgetGrantsAndGifts;
    private double pandLOtherIncome;
    private double pl62145PayrollServiceFees;
    private double pandLExpense;
    private double pandLNetIncome;
    private double pandlProfitVariance;
    private double budgetWorkshops;
    private double budgetWorkshopsCamps;
    private double budgetPayingStudentsFTE;
    private final int MONTH_VARIANCE_COLUMN = 25;
    public CashItemAggregator(HashMap<String, Integer> budgetMap, HashMap<String, Double> pandLMap, int targetMonth)
    {
        this.budgetMap = budgetMap;
        this.pandLMap = pandLMap;
        this.targetMonth = targetMonth;
        extractMapValues();
        System.out.println("(5) Computing Combined Budget Sheet Entries");
    }
    public void extractMapValues()
    {
        pl43400DirectPublicSupport = pandLMap.get("Total 43400 Direct Public Support");
        pl43460ContributedServices = pandLMap.get("43460 Contributed Services");//Non cash item...must be subtracted
        pl47204GrantScholarships = pandLMap.get("Total 47204 Grant Scholarship");
        pl47200ProgramIncome = pandLMap.get("Total 47200 Program Income");
        pl47203LeagueScholarship = pandLMap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
        pl45000Investments = pandLMap.get("Total 45000 Investments");
       // pl46400OtherIncome = pandLMap.get("Total 46400 Other Types of Income");
        pl45000Investments = pandLMap.get("Total 45000 Investments");
        //pandLOtherIncome = pandLMap.get("Total 46400 Other Types of Income");
        pl62000Salaries = pandLMap.get("Total 62000 Salaries & Related Expenses");
        pl62145PayrollServiceFees = pandLMap.get("62145 Payroll Service Fees");
        //budgetGrantsGifts = budgetMap.get("Grants and Gifts");
        budgetMonthlyTuition = budgetMap.get("Monthly Tuition");
        budgetWorkshopsCamps = budgetMap.get("Workshops/Camps");
        budgetSalaries = budgetMap.get("Salaries");
        budgetContractServices = budgetMap.get("Contract Services");
        //budgetRent = budgetMap.get("Rent");
        budgetOperations = budgetMap.get("Operations");
        budgetTotalExpenses = budgetMap.get("Total Expenses");
//        budgetProfit = budgetMap.get("Profit");
        budgetPayingStudentsFTE = budgetMap.get("Paying Students (FTE)");
        pl62100ContractServices = pandLMap.get("Total 62100 Contract Services");
        pl62800Rent = pandLMap.get("Total 62800 Facilities and Equipment");
        //pl68300Travel = pandLMap.get("Total 68300 Travel and Meetings");
        pl65100OtherExpenses = pandLMap.get("Total 65100 Other Types of Expenses");
        //pl65055BreakRoomSupplies = pandLMap.get("65055 Breakroom Supplies");
        pl65000Operations = pandLMap.get("Total 65000 Operations");
//        pl60900BusinessExpenses = pandLMap.get("Total 60900 Business Expenses");
       // pandLMiscExpense = pandLMap.get("Total Other Expenses");
        pandLExpense = pandLMap.get("Total Expenses");
        pandLIncome = pandLMap.get("Total Income");
        pandLNetIncome = pandLMap.get("Net Income");
    }
    public void computeGrantsAndGifts()
    {
        pandLDonations = pl43400DirectPublicSupport - pl43460ContributedServices + pl47204GrantScholarships;
        donationVariance = pandLDonations - budgetGrantsAndGifts;
        printConsoleSummary("Grants And Gifts", budgetGrantsAndGifts, pandLDonations, donationVariance);
    }
    public void computeMonthlyTuition()
    {
        pandLTuition = pl47200ProgramIncome - pl47203LeagueScholarship;
        tuitionVariance = pandLTuition - budgetMonthlyTuition;
        printConsoleSummary("Monthly Tuition", budgetMonthlyTuition, pandLTuition, tuitionVariance);
    }
    public void computeTotalIncome()
    {
        pandLMiscIncome = pl45000Investments + pandLOtherIncome;
        pandLTotalIncome = pandLDonations + pandLTuition + pandLMiscIncome - pl47204GrantScholarships;//grants in both Tuition and Donations
        budgetTotalIncome = budgetGrantsGifts + budgetMonthlyTuition + budgetWorkshopsCamps;
        incomeTotalVariance = pandLTotalIncome - budgetTotalIncome;
        printConsoleSummary("Total Income", budgetTotalIncome, pandLTotalIncome, incomeTotalVariance);
    }
    public void computeSalaries()
    {
        pl62000Salaries = pl62000Salaries + pl62145PayrollServiceFees - pl43460ContributedServices;
        salaryVariance = pl62000Salaries - budgetSalaries;
        printConsoleSummary("Salaries", budgetSalaries, pl62000Salaries, salaryVariance);
    }
    public void computeContractServices()
    {
        contractServiceVariance = pl62100ContractServices - budgetContractServices;
        printConsoleSummary("Contract Services", budgetContractServices, pl62100ContractServices, contractServiceVariance);
    }
    public void computeRent()
    {
        rentVariance = pl62800Rent - budgetRent;
        printConsoleSummary("Rent", budgetRent, pl62800Rent, rentVariance);
    }
    public void computeOperatons()
    {
        pl65000Operations = pl65000Operations + pl65055BreakRoomSupplies + pl65100OtherExpenses + pl68300Travel + pl60900BusinessExpenses;
        operationsVariance = pl65000Operations - budgetOperations;
        printConsoleSummary("Operations", budgetOperations, pl65000Operations, operationsVariance);
    }
    public void computeTotalExpenses()
    {
        pandLTotalExpenses = pl62000Salaries + pl62100ContractServices + pl62800Rent + pl65000Operations + pandLMiscExpense;
        expenseTotalVariance = pandLTotalExpenses - budgetTotalExpenses;
        printConsoleSummary("Total Expenses", budgetTotalExpenses, pandLTotalExpenses, expenseTotalVariance);
    }
    public void computeProfit()
    {
        pandLProfit = pandLTotalIncome - pandLTotalExpenses;
        profitVariance = pandLProfit - budgetProfit;
        printConsoleSummary("Profit", budgetProfit, pandLProfit, profitVariance);
    }
    public void computeStudents()
    {
        actualPayingStudents = (int) (pandLTuition / 240);//Derived...including workshops, slams, etc and partial paying students
        payingStudentsVariance = actualPayingStudents - budgetPayingStudents;
        printConsoleSummary("Paying Students", budgetPayingStudents, actualPayingStudents, payingStudentsVariance);
    }
    public void reconcile()
    {
        pandlProfitVariance = pandLProfit - pandLNetIncome;
        printConsoleSummary("", "", "PROFIT RECONCILIATION", "");
        printConsoleSummary("ACCOUNT", "BUDGET NUMBERS", "P&L NUMBERS", "VARIANCE");//Reconciliation
        printConsoleSummary("------------------------------------", "------------", "------------", "----------");
        System.out.println();
        printConsoleSummary("Profit", pandLProfit, pandLNetIncome, pandlProfitVariance);
        System.out.println("(6) Finished computing Budget Sheet Entries");
    }
    public void printConsoleSummary(String title1, double title2, double title3, double title4)
    {
        System.out.printf("%-40s %,-20.0f %,-20.0f %,-20.0f %n", title1, title2, title3, title4);
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
        printConsoleSummary("ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        printConsoleSummary("------------------------------------", "-------------", "-------------", "---------------------");
        System.out.println();
        LocalDateTime now = LocalDateTime.now();
        XSSFSheet budgetSheet = budgetWorkbook.getSheetAt(0);
        for (Row row : budgetSheet)
        {
            if (row.getRowNum() == 0 || row.getRowNum() == 1)
            {
                row.createCell(MONTH_VARIANCE_COLUMN, XSSFCell.CELL_TYPE_STRING);//For month variance numbers
                try
                {
                    budgetSheet.getRow(0).getCell(0).setCellValue("Updated: " + now);
                    budgetSheet.getRow(0).getCell(MONTH_VARIANCE_COLUMN).setCellValue("Month " + targetMonth);
                    budgetSheet.getRow(1).getCell(MONTH_VARIANCE_COLUMN).setCellValue("VARIANCE");
                    budgetSheet.getRow(1).getCell((int) targetMonth).setCellValue(">>ACTUAL<");
                }
                catch (Exception e)
                {
                    System.out.println("Error updating budget sheet in CashItemAggregator line 238");
                }
            }
            else
            {
                row.createCell(MONTH_VARIANCE_COLUMN, XSSFCell.CELL_TYPE_NUMERIC);//For month variance numbers
                DataFormat format = budgetWorkbook.createDataFormat();
                CellStyle cellstyle = budgetWorkbook.createCellStyle();
                cellstyle.setDataFormat(format.getFormat("#,##0"));
                row.getCell(MONTH_VARIANCE_COLUMN).setCellStyle(cellstyle);
            }
            if (row.getCell(0) != null)
            {
                switch (row.getCell(0).getStringCellValue())
                {
                    case "Grants and Gifts":
                        computeGrantsAndGifts();
                        row.getCell((int) targetMonth).setCellValue(pandLDonations);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) donationVariance);
                        break;
                    case "Monthly Tuition":
                        computeMonthlyTuition();
                        row.getCell((int) targetMonth).setCellValue(pandLTuition);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue(tuitionVariance);
                        break;
                    case "Total Income":
                        computeTotalIncome();
                        row.getCell((int) targetMonth).setCellValue(pandLTotalIncome);
                        row.getCell(13).setCellValue((int) incomeTotalVariance);
                        break;
                    case "Salaries":
                        computeSalaries();
                        row.getCell((int) targetMonth).setCellValue(pl62000Salaries);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) salaryVariance);
                        break;
                    case "Contract Services":
                        computeContractServices();
                        row.getCell((int) targetMonth).setCellValue(pl62100ContractServices);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) contractServiceVariance);
                        break;
                    case "Rent":
                        computeRent();
                        row.getCell((int) targetMonth).setCellValue(pl62800Rent);
                        row.getCell(13).setCellValue((int) rentVariance);
                        break;
                    case "Operations":
                        computeOperatons();
                        row.getCell((int) targetMonth).setCellValue(pl65000Operations);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) operationsVariance);
                        break;
                    case "Total Expenses":
                        computeTotalExpenses();
                        row.getCell((int) targetMonth).setCellValue(pandLTotalExpenses);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) expenseTotalVariance);
                        break;
                    case "Profit":
                        computeProfit();
                        row.getCell((int) targetMonth).setCellValue(pandLProfit);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) profitVariance);
                        break;
                    case "Profit Variance":
                        row.getCell(targetMonth).setCellValue((int) profitVariance);
                        break;
                    case "Paying Students":
                        row.getCell((int) targetMonth).setCellValue(actualPayingStudents);
                        row.getCell(MONTH_VARIANCE_COLUMN).setCellValue((int) payingStudentsVariance);
                        break;
                    default:
                }
            }
        }
        System.out.println("(8) Finished updating budget XSSFsheet");
    }
}
