package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200726
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.time.LocalDate;
import java.util.HashMap;

import static java.lang.Integer.parseInt;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

public class CashItemAggregator
{
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
    private int plContractServices;
    private int vicNetIncome;
    private Cell currentBudgetCell;
    private Cell monthVarianceCell;
    private int vicContractServices;
    private int budgetContractServices;
    private int contractServiceVariance;
    private final int monthVarianceColumnIndex = 13;
    private final int ytdVarianceColumnIndex = 14;
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
    private double budgetFacilities;
    private double payingStudents;
    private double payingStudentsBudget;
    private int payingStudentsDerived;
    private double payingStudentsVariance;
    private String switchKey;
    private int payingStudentsActual;
    private int currentBudgetRowIndex;
    private int currentBudgetColumnIndex;
    private int totalExpenseVariance;
    private int budgetTotalExpenses;
    private Row budgetRow;
    private LocalDate now;
    private int updateHeaderColumnIndex = 0;
    private int header0RowIndex = 0;
    private int header1RowIndex = 1;

    public void aggregateBudget(XSSFWorkbook budgetWorkbook, HashMap<String, Integer> pandLmap, int targetMonth)
    {
        this.budgetWorkbook = budgetWorkbook;
        now = LocalDate.now();
        budgetSheet = budgetWorkbook.getSheetAt(0);
        for (Row row:budgetSheet)//Create variance columns
        {
            row.createCell(monthVarianceColumnIndex,0);//set numeric type
            row.createCell(ytdVarianceColumnIndex,0);//set numeric type
        }
        budgetSheet.getRow(header0RowIndex).getCell(updateHeaderColumnIndex).setCellValue(now + " Update");
        budgetSheet.getRow(header0RowIndex).getCell(monthVarianceColumnIndex).setCellValue("Month " + targetMonth);
        budgetSheet.getRow(header1RowIndex).getCell(monthVarianceColumnIndex).setCellValue("VARIANCE");
        budgetSheet.getRow(header0RowIndex).getCell(ytdVarianceColumnIndex).setCellValue("YTD");
        budgetSheet.getRow(header1RowIndex).getCell(ytdVarianceColumnIndex).setCellValue("VARIANCE");
        budgetSheet.getRow(header1RowIndex).getCell(targetMonth).setCellValue("*Actual");
        System.out.println("Aggregating, Month " + targetMonth + " Budget Proof => ");
        System.out.printf("%-40s %-20s %-20s %-20s %n", "BUDGET ACCOUNT", "BUDGET AMOUNT", "PandL AMOUNT", "Month " + targetMonth + " VARIANCE");
        for (int i = 0; i < budgetSheet.getLastRowNum() - 1; i++)
        {
            budgetRow = budgetSheet.getRow(i);
            currentBudgetCell = budgetRow.getCell(targetMonth);
            currentBudgetRowIndex = budgetRow.getRowNum();
            currentBudgetColumnIndex = targetMonth;
            monthVarianceCell = budgetRow.getCell(monthVarianceColumnIndex);
            try
            {
                switchKey = budgetRow.getCell(0).getStringCellValue().trim();
            }
            catch (Exception e)
            {
                System.out.println("End reading budget rows");;
            }
            switch (switchKey)
            {
                case "Total 43400 Direct Public Support":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    pandlCorporateContributions = pandLmap.get("Total 43400 Direct Public Support");
                    pandlIndividualBusinessContributions = pandLmap.get("43450 Individ, Business Contributions");
                    pandlGrants = pandLmap.get("43455 Grants");
                    contributedServices = pandLmap.get("43460 Contributed Services");
                    investments = pandLmap.get("Total 45000 Investments");
                    totalGrantScholarship = pandLmap.get("Total 47204 Grant Scholarship");
                    vicDirectPublicSupport = pandlCorporateContributions + pandlIndividualBusinessContributions + pandlGrants - contributedServices + investments + totalGrantScholarship;
                    budgetDirectPublicSupport = (int) currentBudgetCell.getNumericCellValue();
                    budgetDirectPublicSupportVariance = vicDirectPublicSupport - budgetDirectPublicSupport;
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Direct Public Support", budgetDirectPublicSupport, vicDirectPublicSupport, budgetDirectPublicSupportVariance);
                    monthVarianceCell.setCellValue(budgetDirectPublicSupportVariance);
                    System.out.println("sw0");
                    break;
                case "Total 47201 Tuition  Fees":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    totalTuitionFees = pandLmap.get("Total 47201 Tuition  Fees");
                    totalWorkshopFees = pandLmap.get("Total 47202 Workshop Fees");
                    vicTuitionFees = totalTuitionFees + totalWorkshopFees;
                    budgetTuitionFeeVariance = (int) (vicTuitionFees - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Tuition  Fees", (int) currentBudgetCell.getNumericCellValue(), vicTuitionFees, budgetTuitionFeeVariance);
                    monthVarianceCell.setCellValue(budgetTuitionFeeVariance);
                    System.out.println("sw1");
                    break;
                case "Total Income":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    vicTotalIncome = vicDirectPublicSupport + vicTuitionFees + totalGrantScholarship;
                    budgetTotalIncome = currentBudgetCell.getNumericCellValue();
                    budgetTotalIncomeVariance = (int) (vicTotalIncome - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Total Income", (int) currentBudgetCell.getNumericCellValue(), vicTotalIncome, budgetTotalIncomeVariance);
                    monthVarianceCell.setCellValue(budgetTotalIncomeVariance);
                    System.out.println("sw2");
                    break;
                case "Total 62000 Salaries & Related Expenses":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    totalSalaries = pandLmap.get("Total 62000 Salaries & Related Expenses");
                    payrollServiceFees = pandLmap.get("62145 Payroll Service Fees");
                    vicSalaries = totalSalaries + payrollServiceFees;
                    budgetTotalSalaries = (int) currentBudgetCell.getNumericCellValue();
                    monthVarianceCell.setCellValue(vicSalaries - budgetTotalSalaries);
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Salaries", (int) currentBudgetCell.getNumericCellValue(), vicSalaries, vicSalaries - totalSalaries);
                    monthVarianceCell.setCellValue(vicSalaries);
                    System.out.println("sw3");
                    break;
                case "Total 62100 Contract Services":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    plContractServices = pandLmap.get("Total 62100 Contract Services");
                    budgetContractServices = (int) currentBudgetCell.getNumericCellValue();
                    contractServiceVariance = plContractServices - budgetContractServices;
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Total Contract Services", (int) currentBudgetCell.getNumericCellValue(), plContractServices, contractServiceVariance);
                    monthVarianceCell.setCellValue(contractServiceVariance);
                    System.out.println("sw4");
                    break;
                case "Total 62800 Facilities and Equipment":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    pandlFacilitiesAndEquipment = pandLmap.get("Total 62800 Facilities and Equipment");
                    pandlDepreciation = pandLmap.get("62810 Depr and Amort - Allowable");
                    vicFacilities = pandlFacilitiesAndEquipment - pandlDepreciation;
                    budgetFacilities = currentBudgetCell.getNumericCellValue();
                    monthVarianceCell.setCellValue(vicFacilities - pandlFacilitiesAndEquipment);
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Facilities and Equipment", (int) currentBudgetCell.getNumericCellValue(), plContractServices, vicFacilities - pandlFacilitiesAndEquipment);
                    System.out.println("sw5");
                    break;
                case "Total 65000 Operations":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    supplies = pandLmap.get("Total 65040 Supplies");
                    operations = pandLmap.get("Total 65000 Operations");
                    pandlTotalExpenses = pandLmap.get("Total 65100 Other Types of Expenses");
                    travel = pandLmap.get("Total 68300 Travel and Meetings");
                    penalties = pandLmap.get("90100 Penalties");
                    vicOperations = supplies + operations + pandlTotalExpenses + travel + penalties;
                    budgetOperationsValue = (int) currentBudgetCell.getNumericCellValue();
                    operationsVarience = (int) (vicOperations - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Operations", (int) currentBudgetCell.getNumericCellValue(), vicOperations, vicOperations - (int) currentBudgetCell.getNumericCellValue());
                    monthVarianceCell.setCellValue(operationsVarience);
                    System.out.println("sw6");
                    break;
                case "Total Expenses":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    vicExpenses = vicSalaries + vicContractServices + vicFacilities + vicOperations;
                    monthVarianceCell.setCellValue(vicExpenses - currentBudgetCell.getNumericCellValue());
                    System.out.println("sw7");
                    break;
                case "Net Income":
                    currentBudgetCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);//numeric
                    vicNetIncome = vicTotalIncome - pandlTotalExpenses;
                    monthVarianceCell.setCellValue(vicNetIncome - currentBudgetCell.getNumericCellValue());
                    System.out.printf("%-40s %-20d %-20d %-20d %n", "Net Cash Income", (int) currentBudgetCell.getNumericCellValue(), pandlTotalExpenses, pandlTotalExpenses - (int) currentBudgetCell.getNumericCellValue());
                    System.out.println("sw8");
                    break;
                case "Paying Students (Actual)":
                    try//Check for void actual student count entry...actual student count not entered in budget yet
                    {
                        payingStudentsActual = (int) currentBudgetCell.getNumericCellValue();
                        System.out.printf("%-40s %-20d %n", "Paying Students (Actual)", (int) currentBudgetCell.getNumericCellValue());
                    }
                    catch (Exception e)
                    {
                        System.out.println("Exception in case sw10: Paying Students (Actual)...probably empty budget cell => " + e);
                    }
                    System.out.println("sw10");
                    break;
                case "Paying Students (Budget)":
                    payingStudentsBudget = 999;
                    payingStudentsVariance = payingStudentsActual - (int)payingStudentsBudget;
                    monthVarianceCell.setCellValue(payingStudentsVariance);
                    System.out.printf("%-40s %-20f %n", "Paying Students (Budget)", payingStudentsBudget);
                    System.out.println("sw9");
                    break;
                default:
                    System.out.println("switch default...SwitchKey => "  + switchKey);
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



