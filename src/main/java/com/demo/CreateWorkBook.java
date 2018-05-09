package com.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class CreateWorkBook {
    public static void main(String[] args)  {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(" Employee Info ");
        XSSFRow row;
        Map< String, Object[] > empinfo = new TreeMap< String, Object[] >();
        empinfo.put( "1", new Object[] {
                "EMP ID", "EMP NAME", "DESIGNATION" });
        empinfo.put( "2", new Object[] {
                "tp01", "Gopal", "Technical Manager" });
        empinfo.put( "3", new Object[] {
                "tp02", "Manisha", "Proof Reader" });
        empinfo.put( "4", new Object[] {
                "tp03", "Masthan", "Technical Writer" });
        empinfo.put( "5", new Object[] {
                "tp04", "Satish", "Technical Writer" });
        empinfo.put( "6", new Object[] {
                "tp05", "Krishna", "Technical Writer" });
        Set< String > keyid = empinfo.keySet();
        int rowid = 0;
        for (String key : keyid)
        {
            row = sheet.createRow(rowid++);
            Object [] objectArr = empinfo.get(key);
            int cellid = 0;
            for (Object obj : objectArr)
            {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
        File file =new File("createworkbook.xlsx");
        try {
            FileOutputStream out =new FileOutputStream(file);
            workbook.write(out);
            out.close();
            System.out.println("createworkbook.xlsx written successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
