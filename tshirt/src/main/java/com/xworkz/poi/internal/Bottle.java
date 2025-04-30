package com.xworkz.poi.internal;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Bottle {

    public void readExcel() throws IOException {
        String path="C:\\Users\\shrih\\OneDrive\\Pictures\\Documents\\Excel-reader-Writer\\tshirt\\src\\main\\resources\\info.xlsx"; //this will set the path to the Excel sheet .//-> refers to current workspace
        FileInputStream fileInputStream=new FileInputStream(path); // this class FileInputStream is used  open the Excel sheet

        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream); // this is to workbook from the file
       // XSSFSheet Sheet = workbook.getSheetAt(0);// this is to get the sheet from the file
        XSSFSheet sheet= workbook.getSheet("sheet1");

        int rows=sheet.getLastRowNum();
        int cols=sheet.getRow(1).getLastCellNum();


        for(int r=0; r<=rows;r++){
            XSSFRow row=sheet.getRow(r);

            for (int c=0;c<cols;c++){
                XSSFCell cell=row.getCell(c);
               double id=0;
               String name;

                switch (cell.getCellType()){
                    case STRING:System.out.println(cell.getStringCellValue()); break;
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue()); break;
                }
            }
        }


    }
}
