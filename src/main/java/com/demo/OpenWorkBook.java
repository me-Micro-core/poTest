package com.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class OpenWorkBook {
    public static void main(String[] args) {
        File file =new File("createworkbook.xlsx");
        try {
            FileInputStream in = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();
            while (rows.hasNext()){
                Row row = rows.next();
                Iterator<Cell> cells = row.iterator();
                while (cells.hasNext()){
                    Cell cell = cells.next();
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(
                                    cell.getNumericCellValue() + " \t\t " );
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(
                                    cell.getStringCellValue() + " \t\t " );
                            break;
                    }
                }
                System.out.println();


            }
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
