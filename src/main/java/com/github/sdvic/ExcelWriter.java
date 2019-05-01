package com.github.sdvic;
/****************************************************************************************
 *  * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.6 April 27, 2019
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class ExcelWriter
{
    public FileOutputStream fileStreamOut;
    public Workbook sarah5yearLocalWorkbook = new XSSFWorkbook();
    public File fiveYearPlanFile;

    public ExcelWriter(Workbook sarah5yearLocalWorkbook, File fiveYearPlanFile)
    {
        this.sarah5yearLocalWorkbook = sarah5yearLocalWorkbook;
        this.fiveYearPlanFile = fiveYearPlanFile;
    }
    public void write5YearPlan()
    {
        try
        {
            System.out.println("..........................................................start writing 5 year plan");
            fileStreamOut = new FileOutputStream(fiveYearPlanFile);
            sarah5yearLocalWorkbook.write(fileStreamOut);
            fileStreamOut.close();
        }
        catch (Exception e)
        {
            System.out.println("Problem writing 5 year plan");
        }
        System.out.println("..........................................................finished writing 5 year plan");
    }
}
