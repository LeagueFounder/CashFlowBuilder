package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L
 * rev 1.1 April 8, 2019
 * copyright 2019 Vic Wintriss 
 ****************************************************************************************/

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Main
{
    public static XSSFWorkbook pandlWorkbook;//P&L
    public static XSSFWorkbook sarah5yearWorkbook;//5Year
    public static Sheet pandlSheet;//P&L
    public static Sheet sarah5yearSheet;//5Year
    public static File pandLfile;
    public static File sarahFiveYearFile;
    public static FileInputStream plfis;
    public static FileInputStream s5yrfis;
    public static FileOutputStream plfos;
    public static FileOutputStream s5yrfos;

    public static void main(String[] args)
    {
        System.out.println("Cash Flow Generator ver 1.1  4/5/19");
        new Setup();
        new ExtractPandLinfoForProjections().cashProjectionSetup();
    }


}
