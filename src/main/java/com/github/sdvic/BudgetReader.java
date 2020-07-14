package com.github.sdvic;
/******************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * version 200713
 * copyright 2020 Vic Wintriss
 ******************************************************************************************/
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

public class BudgetReader
{
    private final File budgetInputFile = new File("/Users/VicMini/Desktop/SarahColorCodedMasterBudget2020.xlsx");
    private  FileInputStream budgetInputFIS;
    private CellStyle backgroundStyle;
    private XSSFWorkbook budgetWorkbook;
    private XSSFWorkbook budgetWorkBook;
    private XSSFRow row;
    private Date now;
    private Calendar cals;

    public void readBudget()
    {
        try
        {
            budgetInputFIS = new FileInputStream(budgetInputFile);
            budgetWorkBook = new XSSFWorkbook(budgetInputFIS);
            budgetInputFIS.close();
        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        cals = Calendar.getInstance();
        now = cals.getTime();
        CellStyle headerCellStyle = budgetWorkBook.getSheetAt(0).getWorkbook().createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        budgetWorkBook.getSheetAt(0).getRow(0).getCell(0).setCellStyle(headerCellStyle);
        budgetWorkBook.getSheetAt(0).getRow(0).getCell(0).setCellValue("Budget Run Date: " + now);
        //budgetSheet.forEach(row -> System.out.println( row.getCell(6) ));
        //cell.setCellStyle(backgroundStyle);
    }

    public XSSFWorkbook getBudgetWorkBook()
    {
        return budgetWorkBook;
    }
}
