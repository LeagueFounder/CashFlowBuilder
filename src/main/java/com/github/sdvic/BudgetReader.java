package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200829
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

public class BudgetReader
{
    String budgetInputFileName = "/Users/VicMini/Desktop/SarahOriginalBudget2020.xlsx";
    String updateInputFileName = "/Users/VicMini/Desktop/Updated2020MasterBudgetOutputFile.xlsx";
    private File budgetInputFile;
    private FormulaEvaluator evaluator;
    private FileInputStream budgetInputFIS;
    private XSSFWorkbook budgetWorkBook;
    private XSSFSheet budgetSheet;
    private HashMap<String, Integer> budgetMap = new HashMap<>();
    private int budgetValue;
    private String budgetKey;
    private XSSFCell cell;
    private String followOnAnswer;

    public void readBudget(int targetMonth, String followOnAnswer)
    {
        System.out.println("(3) Starting reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap, HashMap size: " + budgetMap.size());
        try
        {
            if (followOnAnswer.equals("Yes"))
            {
                budgetInputFile = new File(budgetInputFileName);
            }else
            {
                budgetInputFile = new File(updateInputFileName);
            }
            budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            budgetInputFIS.close();
            budgetSheet = budgetWorkBook.getSheetAt(0);
        }
        catch (FileNotFoundException e)
        {
            System.out.println("file not found");
            e.printStackTrace();
        }
        catch (IOException e)
        {
            System.out.println("file IOexception");
            e.printStackTrace();
        }
        evaluator = budgetWorkBook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        for (Row row : budgetSheet)
        {
                if (row.getCell(0) != null && row.getCell(targetMonth) != null)
                {
                    switch (row.getCell(0).getCellType())//Get budget key
                    {
                        case XSSFCell.CELL_TYPE_BLANK://Type 3
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            break;
                        case XSSFCell.CELL_TYPE_FORMULA://Type 2
                            System.out.println("Found formula...looking for budget key");
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            System.out.println("Found number...looking for budget key");
                            break;
                        case XSSFCell.CELL_TYPE_STRING://Type 1
                            budgetKey = row.getCell(0).getStringCellValue();//Key
                            break;
                        default:
                            System.out.println("switch error");
                    }
                    switch (row.getCell(targetMonth).getCellType())//Get budget value
                    {
                        case XSSFCell.CELL_TYPE_BLANK://Type 3
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            break;
                        case XSSFCell.CELL_TYPE_FORMULA://Type 2
                            budgetValue = (int) row.getCell(targetMonth).getNumericCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            budgetValue = (int) row.getCell(targetMonth).getNumericCellValue();//Value
                            break;
                        case XSSFCell.CELL_TYPE_STRING://Type 1
                            break;
                        default:
                            System.out.println("switch error");
                    }
                    budgetMap.put(budgetKey, budgetValue);
                }
        }
        //budgetMap.forEach((K, V) -> System.out.println( K + " => " + V ));
        System.out.println("(4) Finished reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap, HashMap size: " + budgetMap.size());
    }

    public HashMap<String, Integer> getBudgetMap()
    {
        return budgetMap;
    }

    public XSSFWorkbook getBudgetWorkBook()
    {
        return budgetWorkBook;
    }

    public void setBudgetWorkBook(XSSFWorkbook budgetWorkBook)
    {
        this.budgetWorkBook = budgetWorkBook;
    }
}
