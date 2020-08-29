package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200829
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
    private int pandlProgramIncome;
    private int grantsAndGifts;
    private int pandlOtherIncome;
    private int budgetInvestments;
    private int budgetCashOnlyIncome;
    private int budgetCashTuitionFees;
    private int budgetRent;
    private int cashOnlyExpenses;
    private int pandlTotalGrantScholarship;
    private int pandlOtherExpenses;
    private int pandlLeagueScholarship;
    private int pandlGrantsAndGifts;
    private int pandlTotalProgramIncome;
    private int pandlRent;
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
    private int cashTotalExpenses;
    private int cashProfit;
    private int grantsGiftsVariance;
    private int tuitionVariance;
    private int budgetGrantsGifts;
    private int budgetTuition;
    private int salaryVariance;
    private int rentVariance;
    private int profitVariance;
    private int cashPayingStudents;
    private int budgetPayingStudents;
    private int targetMonth;
    private int pandlBusinessExpenses;
    private int reconcoleIncomeVariance;
    private int reconcileBudgetTotalIncome;
    private int budgetTotalExpense;
    private int pandlTotalExpenses;
    private int expenseTotalVariance;
    private int pandlTotalIncome;
    private int incomeTotalVariance;
    private int budgetMiscIncome;
    private int budgetMiscExpenses;
    private int cashMiscExpenses;
    private int miscExpenseVariance;
    private int cashMiscIncome;
    private int miscIncomeVariance;
    private int cashBusinessExpenses;
    private int MiscExpenseVariance;
    private int pandlNetIncome;

    /******************************************************************************************
     * Compute budget sheet entries
     ******************************************************************************************/
    public void computeCombinedCashBudgetSheetEntries(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, XSSFWorkbook budgetWorkBook, int targetMonth)
    {
        System.out.println("(5) Computing Combined Budget Sheet Entries");
        System.out.printf("%n %-40s %-20s %-20s %-20s %n", "ACCOUNT", "BUDGET AMOUNT", "P&L AMOUNT", "MONTH " + targetMonth + " VARIANCE");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "------------------------------------", "-------------", "-------------", "---------------------");
        try {
          /*************************************************************************************************************
           * GRANTS AND GIFTS
           *************************************************************************************************************/
           pandlGrantsAndGifts = pandLmap.get("Total 43400 Direct Public Support");
           pandlContributedServices = pandLmap.get("43460 Contributed Services");//Non cash item...must be subtracted
           budgetGrantsGifts = budgetMap.get("Grants and Gifts");
           cashGrantsGifts = pandlGrantsAndGifts - pandlContributedServices;
           grantsGiftsVariance = cashGrantsGifts - budgetGrantsGifts;
           System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Grants and Gifts", budgetGrantsGifts, cashGrantsGifts, grantsGiftsVariance);
          /*************************************************************************************************************
           * TUITION
           *************************************************************************************************************/
           pandlProgramIncome = pandLmap.get("Total 47200 Program Income");
           pandlLeagueScholarship = pandLmap.get("Total 47203 League Scholarship");//Non cash item...must be subtracted
           budgetTuition = budgetMap.get("Tuition");
           cashTuition = pandlProgramIncome - pandlLeagueScholarship;
           tuitionVariance = cashTuition - budgetTuition;
           System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition", budgetTuition, cashTuition, tuitionVariance);
          /*************************************************************************************************************
           * MISC INCOME
           *************************************************************************************************************/
           pandlInvestments = pandLmap.get("Total 45000 Investments");
           budgetMiscIncome = budgetMap.get("Misc Income");
           cashMiscIncome = pandlInvestments;
           miscIncomeVariance = cashMiscIncome - budgetMiscIncome;
           System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Misc Income", budgetMiscIncome, cashMiscIncome, miscIncomeVariance);
           /*************************************************************************************************************
            * TOTAL INCOME
            *************************************************************************************************************/
            pandlTotalIncome = pandLmap.get("Total Income");
            cashTotalIncome = cashGrantsGifts + cashTuition;
            budgetTotalIncome = budgetGrantsGifts + budgetTuition + budgetMiscIncome;
            incomeTotalVariance = pandlTotalIncome - budgetTotalIncome;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", budgetTotalIncome, cashTotalIncome, incomeTotalVariance);
            /*************************************************************************************************************
             * SALARIES
             *************************************************************************************************************/
            pandlSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
            pandlContributedServices = pandLmap.get("62010 Salaries contributed services");//Non cash item...must be subtracted
            pandlPayrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
            budgetSalaries = budgetMap.get("Salaries");
            cashSalaries = pandlSalaries + pandlPayrollServiceFees - pandlContributedServices;
            salaryVariance = cashSalaries - budgetSalaries;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", budgetSalaries, cashSalaries, salaryVariance);
            /*************************************************************************************************************
             * CONTRACT SERVICES
             *************************************************************************************************************/
            pandlContractServices = pandLmap.get("Total 62100 Contract Services");
            budgetContractServices = budgetMap.get("Contract Services");
            cashContractServices = pandlContractServices;
            contractServiceVariance = cashContractServices - budgetContractServices;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Contract Services", budgetContractServices, cashContractServices, contractServiceVariance);
            /*************************************************************************************************************
             * RENT
             *************************************************************************************************************/
            pandlRent = pandLmap.get("Total 62800 Facilities and Equipment");
            budgetRent = budgetMap.get("Rent");
            pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");//Non cash item...must be subtracted
            cashRent = pandlRent - pandlDepreciation;
            rentVariance = cashRent - budgetRent;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Rent",  budgetRent, cashRent, rentVariance);
            /*************************************************************************************************************
             * OPERATIONS
             *************************************************************************************************************/
            pandlOperations = pandLmap.get("Total 65000 Operations");
            pandlBreakRoomSupplies = pandLmap.get("65055 Breakroom Supplies");
            pandlOtherExpenses = pandLmap.get("65100 Other Types of Expenses");
            pandlTravel = pandLmap.get("Total 68300 Travel and Meetings");
            budgetOperations = budgetMap.get("Operations");
            cashOperations = pandlOperations + pandlBreakRoomSupplies + pandlOtherExpenses + pandlTravel;
            operationsVariance = cashOperations - budgetOperations;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", budgetOperations, cashOperations, operationsVariance);
            /*************************************************************************************************************
            * MISC EXPENSES
            *************************************************************************************************************/
            pandlBusinessExpenses = pandLmap.get("Total 60900 Business Expenses");
            budgetMiscExpenses = budgetMap.get("Misc Expenses");
            cashMiscExpenses = pandlBusinessExpenses;
            miscExpenseVariance = cashMiscExpenses - budgetMiscExpenses;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Misc Expenses", budgetMiscExpenses, cashMiscExpenses, miscExpenseVariance);
            /*************************************************************************************************************
             * TOTAL EXPENSES
             *************************************************************************************************************/
            budgetTotalExpenses = budgetMap.get("Total Expenses");
            pandlTotalExpenses = pandLmap.get("Total Expenses");
            cashTotalExpenses = cashSalaries + cashContractServices + cashRent + cashOperations;
            expenseTotalVariance = pandlTotalExpenses - budgetTotalExpenses;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Expenses", budgetTotalExpenses, cashTotalExpenses, expenseTotalVariance);
            /*************************************************************************************************************
             * PROFIT
             *************************************************************************************************************/
            budgetProfit = budgetMap.get("Profit");
            profitVariance = budgetMap.get("Profit Variance");
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Profit", budgetProfit, cashProfit, profitVariance);
            cashProfit = cashTotalIncome - cashTotalExpenses;
            profitVariance = cashProfit - budgetProfit;
            /*************************************************************************************************************
             * STUDENTS
             *************************************************************************************************************/
            cashPayingStudents = budgetMap.get("Paying Students (Actual)");
            budgetPayingStudents = budgetMap.get("Paying Students (Budget)");
            payingStudentsVariance = cashPayingStudents - budgetPayingStudents;
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Paying Students", budgetPayingStudents, cashPayingStudents, payingStudentsVariance);
            /*************************************************************************************************************
             * RECONCILE
             *************************************************************************************************************/
            pandlNetIncome = pandLmap.get("Net Income");//Take out in-kind donations!
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Reconcile Income", budgetTotalIncome, pandlTotalIncome, incomeTotalVariance);
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Reconcile Expenses", budgetTotalExpenses, pandlTotalExpenses, expenseTotalVariance);
            System.out.printf("%-40s %,-20d %,-20d %,-20d %n%n", "Reconcile Profit", budgetProfit, pandlNetIncome, profitVariance);
       }
       catch(Exception e)
       {
           System.out.println("\n**    Getting: " + e.getMessage() + " while tryng to read pandlMap/budgetMap in method => computeCombinedCashBudgetSheetEntries(HashMap<String, Integer> pandLmap, HashMap<String, Integer> budgetMap, int targetMonth){}");
       }
        System.out.println("(6) Finished computing Budget Sheet Entries");
    }
    /******************************************************************************************
     * Update Budget Excel Workbook
     ******************************************************************************************/
    public void updateBudgetWorkbook(XSSFWorkbook budgetWorkbook, int targetMonthColumnIndex)
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
                        row.getCell(13).setCellValue(incomeTotalVariance);
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
                        case "Misc Expense Variance":
                        row.getCell(targetMonthColumnIndex).setCellValue(MiscExpenseVariance);
                        row.getCell(13).setCellValue(miscExpenseVariance);
                        break;
                        case "Total Expenses":
                        row.getCell(targetMonthColumnIndex).setCellValue(cashTotalExpenses);
                        row.getCell(13).setCellValue(expenseTotalVariance);
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
        System.out.println("(8) Finished updating budget XSSFsheet");
    }
}



