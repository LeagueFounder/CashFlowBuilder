package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200730
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.LocalDate;
import java.util.HashMap;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

public class CashItemAggregator
{
    private final int monthVarianceColumnIndex = 13;
    private final int ytdVarianceColumnIndex = 14;
    private XSSFWorkbook budgetWorkbook;
    private XSSFSheet budgetSheet;
    private int pandlCorporateContributions;
    private int pandlIndividualBusinessContributions;
    private int pandlGrants;
    private int vicDirectPublicSupport;
    private int totalTuitionFees;
    private int totalWorkshopFees;
    private int vicTuitionFees;
    private int totalSalaries;
    private int payrollServiceFees;
    private int vicSalaries;
    private int contributedServices;
    private int pandlFacilitiesAndEquipment;
    private int pandlDepreciation;
    private int vicFacilities;
    private int supplies;
    private int operations;
    private int pandlTotalExpenses;
    private int travel;
    private int penalties;
    private int vicOperations;
    private int investments;
    private int vicTotalIncome;
    private int vicExpenses;
    private int vicNetIncome;
    private Cell currentBudgetCell;
    private Cell monthVarianceCell;
    private int vicContractServices;
    private int budgetContractServices;
    private int contractServiceVariance;
    private int totalGrantScholarship;
    private int budgetDirectPublicSupport;
    private int budgetDirectPublicSupportVariance;
    private int pandlDirectPublicSupport;
    private int budgetTuitionFeeVariance;
    private double budgetTotalIncome;
    private int budgetTotalIncomeVariance;
    private int budgetTotalSalaries;
    private int budgetOperationsValue;
    private int operationsVarience;
    private int budgetFacilities;
    private int payingStudents;
    private int payingStudentsBudget;
    private int payingStudentsDerived;
    private int payingStudentsVariance;
    private String switchKey;
    public int payingStudentsActual;
    private int currentBudgetRowIndex;
    private int currentBudgetColumnIndex;
    private int totalExpenseVariance;
    private int budgetTotalExpenses;
    private Row budgetRow;
    private LocalDate now;
    private int updateHeaderColumnIndex = 0;
    private int header0RowIndex = 0;
    private int header1RowIndex = 1;
    private int facilitiesVariance;
    private int expenseVariation;
    private int monthBudgetIncome;
    private int monthIncomeVariance;
    //CELL_TYPE_NUMERIC
    //CELL_TYPE_STRING

    public void aggregateBudget(XSSFWorkbook budgetWorkbook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkbook = budgetWorkbook;
        now = LocalDate.now();
        budgetSheet = budgetWorkbook.getSheetAt(0);
        for (Row row : budgetSheet)//Create variance columns
        {
            row.createCell(monthVarianceColumnIndex, 0);//set numeric type
            row.createCell(ytdVarianceColumnIndex, 0);//set numeric type
        }
        budgetSheet.getRow(header0RowIndex).getCell(updateHeaderColumnIndex).setCellValue(now + " Update");
        budgetSheet.getRow(header0RowIndex).getCell(monthVarianceColumnIndex).setCellValue("Month " + targetMonth);
        budgetSheet.getRow(header1RowIndex).getCell(monthVarianceColumnIndex).setCellValue("VARIANCE");
        budgetSheet.getRow(header0RowIndex).getCell(ytdVarianceColumnIndex).setCellValue("YTD");
        budgetSheet.getRow(header1RowIndex).getCell(ytdVarianceColumnIndex).setCellValue("VARIANCE");
        budgetSheet.getRow(header1RowIndex).getCell(targetMonth).setCellValue("*Actual");
        System.out.println("Aggregating, Month " + targetMonth + " Budget Proof => ");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "BUDGET ACCOUNT", "BUDGET AMOUNT", "Actual AMOUNT", "Month " + targetMonth + " VARIANCE");
        for (int i = 0; i < budgetSheet.getLastRowNum() - 1; i++)
        {
            if(budgetSheet.getRow(i).getCell(0).getStringCellValue().equals("End of Budget"))
            {
                System.out.println("Breaking budgetSheet loop at i = " + i);
                break;
            }
            budgetRow = budgetSheet.getRow(i);
            currentBudgetCell = budgetRow.getCell(targetMonth);
            currentBudgetRowIndex = budgetRow.getRowNum();
            currentBudgetColumnIndex = targetMonth;
            monthVarianceCell = budgetRow.getCell(monthVarianceColumnIndex);
            switchKey = budgetRow.getCell(0).getStringCellValue().trim();
            switch (switchKey)
            {
                case "Total 43400 Direct Public Support":
                    pandlCorporateContributions = pandLmap.get("Total 43400 Direct Public Support");
                    pandlIndividualBusinessContributions = pandLmap.get("43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("43455 Grants");
                    contributedServices = pandLmap.get("43460 Contributed Services");
                    investments = pandLmap.get("Total 45000 Investments");
                    totalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contributedServices + investments + totalGrantScholarship;
                    budgetDirectPublicSupport = (int) currentBudgetCell.getNumericCellValue();
                    budgetDirectPublicSupportVariance = vicDirectPublicSupport - budgetDirectPublicSupport;
                    monthVarianceCell.setCellValue(budgetDirectPublicSupportVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Direct Public Support", budgetDirectPublicSupport, vicDirectPublicSupport, budgetDirectPublicSupportVariance);
                    currentBudgetCell.setCellValue(vicDirectPublicSupport);
                    break;
                case "Total 47201 Tuition  Fees":
                    totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                    budgetTuitionFeeVariance = (int) (vicTuitionFees - currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(budgetTuitionFeeVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition  Fees", (int) currentBudgetCell.getNumericCellValue(), vicTuitionFees, budgetTuitionFeeVariance);
                    currentBudgetCell.setCellValue(vicTuitionFees);
                    break;
                case "Total Income":
                    vicTotalIncome = vicDirectPublicSupport + vicTuitionFees + totalGrantScholarship;
                    budgetTotalIncome = currentBudgetCell.getNumericCellValue();
                    budgetTotalIncomeVariance = (int) (vicTotalIncome - currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(budgetTotalIncomeVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", (int) currentBudgetCell.getNumericCellValue(), vicTotalIncome, budgetTotalIncomeVariance);
                   currentBudgetCell.setCellValue(vicTotalIncome);
                    break;
                case "Total 62000 Salaries & Related Expenses":
                    totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    vicSalaries = totalSalaries + payrollServiceFees;
                    budgetTotalSalaries = (int) currentBudgetCell.getNumericCellValue();
                    monthVarianceCell.setCellValue(vicSalaries - budgetTotalSalaries);
                    monthVarianceCell.setCellValue(vicSalaries);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", (int) currentBudgetCell.getNumericCellValue(), vicSalaries, vicSalaries - totalSalaries);
                    currentBudgetCell.setCellValue(vicSalaries);
                    break;
                case "Total 62100 Contract Services":
                    vicContractServices = pandLmap.get("Total 62100 Contract Services");
                    budgetContractServices = (int) currentBudgetCell.getNumericCellValue();
                    contractServiceVariance = vicContractServices - budgetContractServices;
                    monthVarianceCell.setCellValue(contractServiceVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Contract Services", (int) currentBudgetCell.getNumericCellValue(), vicContractServices, contractServiceVariance);
                   currentBudgetCell.setCellValue(vicContractServices);
                    break;
                case "Total 62800 Facilities and Equipment":
                    pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    vicFacilities = pandlFacilitiesAndEquipment - pandlDepreciation;
                    budgetFacilities = (int) currentBudgetCell.getNumericCellValue();
                    facilitiesVariance = vicFacilities - budgetFacilities;
                    monthVarianceCell.setCellValue(facilitiesVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Facilities and Equipment", (int) currentBudgetCell.getNumericCellValue(), vicFacilities, facilitiesVariance);
                    currentBudgetCell.setCellValue(vicFacilities);
                    break;
                case "Total 65000 Operations":
                    supplies = pandLmap.get("Total 65040 Supplies");
                    operations = pandLmap.get("Total 65000 Operations");
                    pandlTotalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    vicOperations = supplies + operations + pandlTotalExpenses + travel + penalties;
                    budgetOperationsValue = (int) currentBudgetCell.getNumericCellValue();
                    operationsVarience = (int) (vicOperations - currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(operationsVarience);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", (int) currentBudgetCell.getNumericCellValue(), vicOperations, vicOperations - (int) currentBudgetCell.getNumericCellValue());
                    currentBudgetCell.setCellValue(vicOperations);
                    break;
                case "Total Expenses":
                    vicExpenses = vicSalaries + vicContractServices + vicFacilities + vicOperations;
                    totalExpenseVariance = vicExpenses - budgetTotalExpenses;
                    monthVarianceCell.setCellValue(totalExpenseVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Expenses", (int) currentBudgetCell.getNumericCellValue(), vicExpenses, totalExpenseVariance);
                   currentBudgetCell.setCellValue(totalExpenseVariance);
                    break;
                case "Net Income":
                    vicNetIncome = vicTotalIncome - vicExpenses;
                    monthBudgetIncome = (int) currentBudgetCell.getNumericCellValue();
                    monthIncomeVariance = vicNetIncome - monthBudgetIncome;
                    monthVarianceCell.setCellValue(monthIncomeVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Net Cash Income", (int) currentBudgetCell.getNumericCellValue(), vicNetIncome, monthIncomeVariance);
                    currentBudgetCell.setCellValue(vicNetIncome);
                    break;
                case "Paying Students (Actual)":
                    payingStudentsActual = (int) currentBudgetCell.getNumericCellValue();
                    System.out.printf("%-40s %,-40d %n", "Paying Students (Actual)", payingStudentsActual);
                    break;
                case "Paying Students (Budget)":
                    payingStudentsBudget = (int) currentBudgetCell.getNumericCellValue();
                    payingStudentsVariance = payingStudentsActual - payingStudentsBudget;
                    monthVarianceCell.setCellValue(payingStudentsVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Paying Students (Budget)", payingStudentsBudget, payingStudentsActual, payingStudentsVariance);
                    currentBudgetCell.setCellValue(payingStudentsActual);
                    break;
                default:
            }
        }
    }

    public XSSFWorkbook getBudgetWorkbook()
    {
        return budgetWorkbook;
    }

//    public int computeYTDvariance(Row budgetItemRow)
//    {
//        int ytdVariance = 0;
//        int i = 1;
//        while (budgetSheet.getRow(1).getCell(i).getStringCellValue().equals("Actual"))
//        {
//            if (budgetItemRow.getCell(i).getCellType() == CELL_TYPE_NUMERIC)
//            {
//                ytdVariance += budgetItemRow.getCell(i).getNumericCellValue();
//            }
//            i++;
//        }
//        return ytdVariance;
//    }
}



