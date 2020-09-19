package com.github.sdvic;
//******************************************************************************************
// * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
// * version 200918
// * copyright 2020 Vic Wintriss
//******************************************************************************************

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

public class BudgetReader {
    String budgetInputFileName = "/Users/vicwintriss/Desktop/SarahOriginalBudget2020.xlsx";
    String updateInputFileName = "/Users/vicwintriss/Desktop//Updated2020MasterBudgetOutputFile.xlsx";
    private XSSFWorkbook budgetWorkBook;
    private final HashMap<String, Integer> budgetMap = new HashMap<>();
    private XSSFCell cell;
    private XSSFCell keyCell;
    private XSSFCell valueCell;
    private String followOnAnswer;

    public void readBudget(int targetMonth, String followOnAnswer) {
        System.out.println("(3) Starting reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap, HashMap size: " + budgetMap.size());
        XSSFSheet budgetSheet;
        try {
            File budgetInputFile;
            if (followOnAnswer.equals("Yes")) {
                budgetInputFile = new File(budgetInputFileName);
            } else {
                budgetInputFile = new File(updateInputFileName);
            }
            FileInputStream budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            budgetInputFIS.close();
            budgetSheet = budgetWorkBook.getSheetAt(0);
        } catch (FileNotFoundException e) {
            System.out.println("file not found");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("file IOexception");
            e.printStackTrace();
        }
        FormulaEvaluator evaluator = budgetWorkBook.getCreationHelper().createFormulaEvaluator();
        XSSFFormulaEvaluator.evaluateAllFormulaCells(budgetWorkBook);
        budgetSheet = budgetWorkBook.getSheetAt(0);
        for (int rowIndex = 0; rowIndex < budgetSheet.getLastRowNum(); rowIndex++) {
            XSSFRow row = budgetSheet.getRow(rowIndex);
            if (row != null) {
                if (row.getCell(targetMonth) == null)
                {
                    row.createCell(targetMonth);
                    row.getCell(targetMonth).setCellValue(0);
                }
                if (row.getCell(0) != null && row.getCell(targetMonth) != null) {
                    Cell keyCell = row.getCell(0);
                    Cell valueCell = row.getCell(targetMonth);
                    if (keyCell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        if (valueCell.getCellType() == XSSFCell.CELL_TYPE_FORMULA || valueCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                            String keyValue = keyCell.getStringCellValue().trim();
                            int cellValue = (int) valueCell.getNumericCellValue();
                            budgetMap.put(keyValue, cellValue);
                        }
                    }
                }
            }
        }
        //budgetMap.forEach((K, V) -> System.out.println(K + " => " + V));
        System.out.println("(4) Finished reading Budget In budgetReader from " + budgetInputFileName + " to: budgetHashMap, HashMap size: " + budgetMap.size());
    }

    public HashMap<String, Integer> getBudgetMap() {
        return budgetMap;
    }

    public XSSFWorkbook getBudgetWorkBook() {
        return budgetWorkBook;
    }

    public void setBudgetWorkBook(XSSFWorkbook budgetWorkBook) {
        this.budgetWorkBook = budgetWorkBook;
    }
}