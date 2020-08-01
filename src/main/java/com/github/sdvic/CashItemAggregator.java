package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200731
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
    private XSSFWorkbook budgetWorkbook;
    private XSSFSheet budgetSheet;
    private int vicDirectPublicSupport;
    private int vicTuitionFees;
    private int vicSalaries;
    private int vicFacilities;
    private int vicOperations;
    private int vicTotalIncome;
    private int vicExpenses;
    private int vicContractServices;
    private int totalGrantScholarship;
    public int payingStudentsActual;
    private int budgetTotalExpenses;
    //CELL_TYPE_NUMERIC
    //CELL_TYPE_STRING

    public void aggregateBudget(XSSFWorkbook budgetWorkbook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkbook = budgetWorkbook;
        LocalDate now = LocalDate.now();
        budgetSheet = budgetWorkbook.getSheetAt(0);
        int monthVarianceColumnIndex = 13;
        int ytdVarianceColumnIndex = 14;
        for (Row row : budgetSheet)//Create variance columns
        {
            row.createCell(monthVarianceColumnIndex, 0);//set numeric type
            row.createCell(ytdVarianceColumnIndex, 0);//set numeric type
        }
        int updateHeaderColumnIndex = 0;
        int header0RowIndex = 0;
        budgetSheet.getRow(header0RowIndex).getCell(updateHeaderColumnIndex).setCellValue(now + " Update");
        budgetSheet.getRow(header0RowIndex).getCell(monthVarianceColumnIndex).setCellValue("Month " + targetMonth);
        int header1RowIndex = 1;
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
            Row budgetRow = budgetSheet.getRow(i);
            Cell currentBudgetCell = budgetRow.getCell(targetMonth);
            int currentBudgetRowIndex = budgetRow.getRowNum();
            Cell monthVarianceCell = budgetRow.getCell(monthVarianceColumnIndex);
            String switchKey = budgetRow.getCell(0).getStringCellValue().trim();
            switch (switchKey)
            {
                case "Total 43400 Direct Public Support":
                    int pandlCorporateContributions = pandLmap.get("Total 43400 Direct Public Support");
                    int pandlIndividualBusinessContributions = pandLmap.get("43450 Individ, Business Contributions");
                    int pandlGrants = pandLmap.get("43455 Grants");
                    int contributedServices = pandLmap.get("43460 Contributed Services");
                    int investments = pandLmap.get("Total 45000 Investments");
                    totalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contributedServices + investments + totalGrantScholarship;
                    int budgetDirectPublicSupport = (int) currentBudgetCell.getNumericCellValue();
                    int budgetDirectPublicSupportVariance = vicDirectPublicSupport - budgetDirectPublicSupport;
                    monthVarianceCell.setCellValue(budgetDirectPublicSupportVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Direct Public Support", budgetDirectPublicSupport, vicDirectPublicSupport, budgetDirectPublicSupportVariance);
                    currentBudgetCell.setCellValue(vicDirectPublicSupport);
                    break;
                case "Total 47201 Tuition  Fees":
                    int totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    int totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                    int budgetTuitionFeeVariance = (int) (vicTuitionFees - currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(budgetTuitionFeeVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Tuition  Fees", (int) currentBudgetCell.getNumericCellValue(), vicTuitionFees, budgetTuitionFeeVariance);
                    currentBudgetCell.setCellValue(vicTuitionFees);
                    break;
                case "Total Income":
                    vicTotalIncome = vicDirectPublicSupport + vicTuitionFees + totalGrantScholarship;
                    double budgetTotalIncome = currentBudgetCell.getNumericCellValue();
                    int budgetTotalIncomeVariance = (int) (vicTotalIncome - currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(budgetTotalIncomeVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Income", (int) currentBudgetCell.getNumericCellValue(), vicTotalIncome, budgetTotalIncomeVariance);
                   currentBudgetCell.setCellValue(vicTotalIncome);
                    break;
                case "Total 62000 Salaries & Related Expenses":
                    int totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    int payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    vicSalaries = totalSalaries + payrollServiceFees;
                    int budgetTotalSalaries = (int) currentBudgetCell.getNumericCellValue();
                    int monthSalaryVariance = vicSalaries - budgetTotalSalaries;
                    monthVarianceCell.setCellValue(monthSalaryVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Salaries", (int) currentBudgetCell.getNumericCellValue(), vicSalaries, monthSalaryVariance);
                    currentBudgetCell.setCellValue(vicSalaries);
                    break;
                case "Total 62100 Contract Services":
                    vicContractServices = pandLmap.get("Total 62100 Contract Services");
                    int budgetContractServices = (int) currentBudgetCell.getNumericCellValue();
                    int contractServiceVariance = vicContractServices - budgetContractServices;
                    monthVarianceCell.setCellValue(contractServiceVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Total Contract Services", (int) currentBudgetCell.getNumericCellValue(), vicContractServices, contractServiceVariance);
                   currentBudgetCell.setCellValue(vicContractServices);
                    break;
                case "Total 62800 Facilities and Equipment":
                    int pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    int pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    vicFacilities = pandlFacilitiesAndEquipment - pandlDepreciation;
                    int budgetFacilities = (int) currentBudgetCell.getNumericCellValue();
                    int facilitiesVariance = vicFacilities - budgetFacilities;
                    monthVarianceCell.setCellValue(facilitiesVariance);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Facilities and Equipment", (int) currentBudgetCell.getNumericCellValue(), vicFacilities, facilitiesVariance);
                    currentBudgetCell.setCellValue(vicFacilities);
                    break;
                case "Total 65000 Operations":
                    int supplies = pandLmap.get("Total 65040 Supplies");
                    int operations = pandLmap.get("Total 65000 Operations");
                    int pandlTotalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    int travel = pandLmap.get("Total 68300 Travel and Meetings");
                    int penalties = pandLmap.get("90100 Penalties");
                    vicOperations = supplies + operations + pandlTotalExpenses + travel + penalties;
                    int budgetOperationsValue = (int) currentBudgetCell.getNumericCellValue();
                    int operationsVarience = (int) (vicOperations - currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(operationsVarience);
                    System.out.printf("%-40s %,-20d %,-20d %,-20d %n", "Operations", (int) currentBudgetCell.getNumericCellValue(), vicOperations, vicOperations - (int) currentBudgetCell.getNumericCellValue());
                    currentBudgetCell.setCellValue(vicOperations);
                    break;
                case "Total Expenses":
                    vicExpenses = vicSalaries + vicContractServices + vicFacilities + vicOperations;
                    budgetTotalExpenses = (int) currentBudgetCell.getNumericCellValue();
                    int totalExpenseVariance = vicExpenses - budgetTotalExpenses;
                    monthVarianceCell.setCellValue(totalExpenseVariance);
                    printDataProof("Total Expenses", (int) currentBudgetCell.getNumericCellValue(), vicExpenses, totalExpenseVariance);
                    currentBudgetCell.setCellValue(vicExpenses);
                    break;
                case "Net Income":
                    int vicNetIncome = vicTotalIncome - vicExpenses;
                    int monthBudgetIncome = (int) currentBudgetCell.getNumericCellValue();
                    int monthIncomeVariance = vicNetIncome - monthBudgetIncome;
                    monthVarianceCell.setCellValue(monthIncomeVariance);
                    printDataProof( "Net Cash Income", (int) currentBudgetCell.getNumericCellValue(), vicNetIncome, monthIncomeVariance);
                    currentBudgetCell.setCellValue(vicNetIncome);
                    break;
                case "Paying Students (Actual)":
                    payingStudentsActual = (int) currentBudgetCell.getNumericCellValue();
                                        System.out.printf("%-40s %,-40d %n", "Paying Students (Actual)", payingStudentsActual);
                    break;
                case "Paying Students (Budget)":
                    int payingStudentsBudget = (int) currentBudgetCell.getNumericCellValue();
                    int payingStudentsVariance = payingStudentsActual - payingStudentsBudget;
                    monthVarianceCell.setCellValue(payingStudentsVariance);
                    printDataProof("Paying Students (Budget)", payingStudentsBudget, payingStudentsActual, payingStudentsVariance);
                    break;
                default:
            }
        }
    }

    public XSSFWorkbook getBudgetWorkbook()
    {
        return budgetWorkbook;
    }

    public int computeYTDvariance(Row budgetItemRow)
    {
        int ytdVariance = 0;
        int i = 1;
        while (budgetSheet.getRow(1).getCell(i).getStringCellValue().equals("Actual"))
        {
            if (budgetItemRow.getCell(i).getCellType() == CELL_TYPE_NUMERIC)
            {
                ytdVariance += budgetItemRow.getCell(i).getNumericCellValue();
            }
            i++;
        }
        return ytdVariance;
    }
    public void printDataProof(String item, int budgetValue, int actualValue, int variance)
    {
        System.out.printf("%-40s %,-20d %,-20d %,-20d %n", item, budgetValue, actualValue, variance);
    }
}



