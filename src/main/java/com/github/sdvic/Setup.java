package com.github.sdvic;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;

import static com.github.sdvic.Main.*;

public class Setup
{
    public Setup()
    {
        System.out.println("========================== Setting Up =====================");
        try
        {
            pandLfile = new File("/Users/VicMini/Desktop/PandL2018.xlsx");
        }
        catch (Exception e)
        {
            System.out.println("Can't make pandLfile");
        }
        try
        {
            sarahFiveYearFile = new File("/Users/VicMini/Desktop/sarah5yearPlan.xlsx");
        }
        catch(Exception e)
        {
            System.out.println("Can't make sarahFiveYearFile" + e);
        }
        try
        {
            s5yrfis = new FileInputStream(sarahFiveYearFile);
        }
        catch (Exception e)
        {
            System.out.println("Can't get s5yrfis " + e);//**********************
        }
        try
        {
            plfis = new FileInputStream(pandLfile);
        }
        catch (Exception e)
        {
            System.out.println("Can't get plfis " + e);//**********************
        }
        try
        {
            pandlWorkbook = (XSSFWorkbook) WorkbookFactory.create(plfis);
        }
        catch(Exception e)
        {
            System.out.println("Can't create pandLWorkbook");
        }
        try
        {
            pandlSheet = pandlWorkbook.getSheetAt(0);//P&L
        }
        catch (Exception e)
        {
            System.out.println("Can't get pandLSheet " + e);//**********************
        }
        try
        {
            sarah5yearWorkbook = (XSSFWorkbook) WorkbookFactory.create(s5yrfis);
        }
        catch (Exception e)
        {
            System.out.println("Can't get sarah5yearWorkbook " + e);//**********************
        }
        try
        {
            sarah5yearSheet = sarah5yearWorkbook.getSheetAt(0);
        }
        catch (Exception e)
        {
            System.out.println("Can't get sarah5yearSheet " + e);//**********************
        }
    }
}
