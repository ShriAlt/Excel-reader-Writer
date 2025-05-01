package com.xworkz.poi.internal;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Pillow {

    public void  readSheet() throws IOException {

        String excelPath="C:\\Users\\shrih\\OneDrive\\Pictures\\Documents\\Excel-reader-Writer\\tshirt\\src\\main\\resources\\info.xlsx";
        FileInputStream fileInputStream=new FileInputStream(excelPath);

        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
//        XSSFSheet sheet = workbook.getSheet("sheet1");
        XSSFSheet sheet=workbook.getSheetAt(0);

        int rows= sheet.getLastRowNum();
        int cols=sheet.getRow(0).getLastCellNum();


        for ( int r=0; r<=rows; r++){

            XSSFRow row= sheet.getRow(r);
            for( int c=0; c<cols;c++){
                XSSFCell cell=row.getCell(c);

                switch (cell.getCellType()){
                    case STRING:
                        System.out.println(cell.getStringCellValue() + "\t");break;
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue() + "\t");break;
                }

            }System.out.println();
        }
        workbook.close();
        fileInputStream.close();



    }
}
