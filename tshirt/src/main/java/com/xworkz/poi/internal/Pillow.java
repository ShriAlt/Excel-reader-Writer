package com.xworkz.poi.internal;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Pillow {

    public void  readSheet() throws IOException {

        String excelPath="C:\\Users\\shrih\\OneDrive\\Pictures\\Documents\\Excel-reader-Writer\\tshirt\\src\\main\\resources\\sample_excel.xlsx";
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
    public void excelWrite() throws IOException {
        XSSFWorkbook  workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("harsha");
        sheet.createRow(0);
        sheet.getRow(0).createCell(0).setCellValue("name");
        sheet.getRow(0).createCell(1).setCellValue("age");

        sheet.createRow(1);
        sheet.getRow(0).createCell(0).setCellValue("harsha");
        sheet.getRow(0).createCell(1).setCellValue("22");

        sheet.createRow(3);
        sheet.getRow(0).createCell(0).setCellValue("balu");
        sheet.getRow(0).createCell(1).setCellValue("22");

        String excelPath="C:\\Users\\shrih\\OneDrive\\Pictures\\Documents\\Excel-reader-Writer\\tshirt\\src\\main\\resources\\sample_excel1.xlsx";


        FileOutputStream fileOutputStream =new FileOutputStream(excelPath);
//        File file =new File("C:\\Users\\shrih\\OneDrive\\Pictures\\Documents\\Excel-reader-Writer\\tshirt\\src\\main\\resources\\harsha.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
    }


}
