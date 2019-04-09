package com.github.sdvic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.util.ArrayList;
import static com.github.sdvic.Main.pandlSheet;

public class ExtractPandLinfoForProjections
{
    public ArrayList<String> pandLAccountList = new ArrayList<String>();
    public ArrayList<Integer> accountValues = new ArrayList<Integer>();
    public Cell foundCell;

    public void cashProjectionSetup()
    {
        pandLAccountList.add("43450 Individ, Business Contributions");
        pandLAccountList.add("Total 47201 Tuition  Fees");
        pandLAccountList.add("Total 47202 Workshop Fees");
        pandLAccountList.add("Total 47204 Grant Scholarship");
        pandLAccountList.add("62015 Salaries");
        pandLAccountList.add("62041 Health Insurance");
        pandLAccountList.add("62042 Workers Compensation Insurance");
        pandLAccountList.add("62051 Social Security/Medicare");
        pandLAccountList.add("62052 Federal Unemployment Tax");
        pandLAccountList.add("62053 CA SUI");
        pandLAccountList.add("62054 Telephone Reimbursement");
        pandLAccountList.add("62055 Pension Expense");
        pandLAccountList.add("62155 403(b) Plan Fees");
        pandLAccountList.add("62850 Repairs and Maintenance");
        pandLAccountList.add("62870 Property Insurance");
        pandLAccountList.add("62890 Rent");
        pandLAccountList.add("62891 Utilities");
        pandLAccountList.add("62892 San Marcos School District");
        pandLAccountList.add("Total 65000 Operations");
        pandLAccountList.add("62110 Accounting Fees");
        pandLAccountList.add("62155 403(b) Plan Fees");
        pandLAccountList.add("62145 Payroll Service Fees");
        pandLAccountList.add("62800 Facilities and Equipment");
        pandLAccountList.add("65055 Breakroom Supplies");
        pandLAccountList.add("Total 65100 Other Types of Expenses");
        pandLAccountList.add("Total 68300 Travel and Meetings");
        for (String s: pandLAccountList)
        {
            foundCell = findExcelTextEntry(pandlSheet, s);
            int ri = foundCell.getRowIndex();
            int pandlAccountValue = (int)pandlSheet.getRow(ri).getCell(1).getNumericCellValue();
            accountValues.add(pandlAccountValue);
            System.out.println(s + " => " + pandlAccountValue);
        }
    }
    public static Cell findExcelTextEntry(Sheet sheet, String searchName)
    {
        for (Row row : sheet)
        {
            for (Cell cell : row)
            {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING)
                {
                    if (cell.getStringCellValue().contains(searchName))
                    {
                        return cell;
                    }
                }
            }
        }
        Cell cell = pandlSheet.getRow(0).getCell(0);
        return cell;//error...can't find cell name
    }
}
