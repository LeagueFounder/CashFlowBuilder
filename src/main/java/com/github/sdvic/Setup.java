package com.github.sdvic;
/****************************************************************************************
 * Application to extract Cash Flow data from Quick Books P&L and build Cash Projections
 * rev 1.3 April 15, 2019
 * copyright 2019 Vic Wintriss
 ****************************************************************************************/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;

public class Setup
{
    public HashMap<String, Integer> chartOfAccounts = new HashMap<>();
    public Sheet pandlSheet;
    public Sheet sarah5yearLocalSheet;

    public Setup(Sheet pls, Sheet s5Ls)
    {
        System.out.println("Setting Up ==============");
        String cellKey = null;
        int cellValue = 0;
        String title = null;
        int amount = 0;
        pandlSheet = pls;
        sarah5yearLocalSheet = s5Ls;

        System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Reading Chart Of Accounts &&&&&&&&&&&&&&&&&&&&&&&&&");
        for (Row row : pandlSheet)//Bring full chart of accounts from Excel (QuickBooks) P&L into HashMap chartOfAcocounts
        {
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_BLANK)
            {
                continue;
            }
            if (row.getCell(0).getCellType() == Cell.CELL_TYPE_STRING)
            {
                cellKey = row.getCell(0).getStringCellValue();
                System.out.print(cellKey + "      ");
            }
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC)
            {
                cellValue = (int) row.getCell(1).getNumericCellValue();
                System.out.println(cellValue);
            }
            chartOfAccounts.put(cellKey, cellValue);
        }

        System.out.println("%%%%%%%%%%%%%%%%%%%%%%%%%%% Cash Flow Analysis %%%%%%%%%%%%%%%%%%%%%%%");
        for (Row row : sarah5yearLocalSheet)//Bring full cash flow chart from Excel
        {
            if (row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
            {
                continue;
            }
            if (row.getCell(0).getCellType() == Cell.CELL_TYPE_STRING)
            {
                title = row.getCell(0).getStringCellValue();
                System.out.print(title + "         ");
            }
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC)
            {
                amount = (int) row.getCell(1).getNumericCellValue();
                System.out.println(amount + "\n");
            }
            if (row.getCell(1).getCellType() == Cell.CELL_TYPE_BLANK)
            {
                continue;
            }
            System.out.println();
        }
            System.out.println("Chart Of Accounts Loaded=========(" + chartOfAccounts.size() + " entries)");
            System.out.println("Aggregating Cash Contributions****************");
            Integer indBusContributions = chartOfAccounts.get("      43450 Individ, Business Contributions");
            System.out.println("Individual, Business Contributions " + indBusContributions);
            System.out.println("...............................................Total Cash Contributions => " + indBusContributions);
            System.out.println("Aggregating Cash Grants******************");
            Integer javaGrantTuition = chartOfAccounts.get("         47260 Java Tuition- Grant Scholarship");
            System.out.println("Grant Java " + javaGrantTuition);
            Integer grantWorkshops = chartOfAccounts.get("         47280 Java Workshop - Grant Scholarship");
            System.out.println("Grant Workshops " + grantWorkshops);
            int totalGrants = javaGrantTuition + grantWorkshops;
            System.out.println("...............................................Total Cash Grants => " + totalGrants);
            System.out.println("Aggregating Cash Tuition*********************");
            Integer javaTuition = chartOfAccounts.get("      Total 47201 Tuition  Fees");
            System.out.println("Java Tuition " + javaTuition);
            Integer workshopTuition = chartOfAccounts.get("      Total 47202 Workshop Fees");
            System.out.println("Workshop Tuition " + workshopTuition);
            int totalTuition = javaTuition + workshopTuition;
            System.out.println("...............................................Total Cash Tuition => " + totalTuition);
            System.out.println("TOTAL CASH INCOME => " + (totalGrants + totalTuition));
        }
    }
