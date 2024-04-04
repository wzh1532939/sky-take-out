package com.sky.test;

import org.apache.poi.hssf.record.DVALRecord;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;


public class poi {
    public static void todo() throws IOException {
        XSSFWorkbook excel = new XSSFWorkbook();
        XSSFSheet sheet = excel.createSheet("info");
        XSSFRow row = sheet.createRow(1);
        row.createCell(1).setCellValue("姓名");
        row.createCell(2).setCellValue("城市");
        row = sheet.createRow(2);
        row.createCell(1).setCellValue("张三");
        row.createCell(2).setCellValue("北京");
        row = sheet.createRow(3);
        row.createCell(1).setCellValue("李四");
        row.createCell(2).setCellValue("南京");

        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\info.xlsx"));
        excel.write(fileOutputStream);
        fileOutputStream.close();
        excel.close();
    }

    public static void read() throws IOException {
        XSSFWorkbook excel = new XSSFWorkbook(new FileInputStream(new File("D:\\info.xlsx")));
        XSSFSheet sheet = excel.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i <=lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            String cellValue = row.getCell(1).getStringCellValue();
            String cellValue1 = row.getCell(2).getStringCellValue();
            System.out.println(cellValue+" "+cellValue1);
        }
        excel.close();
    }
    public static void main(String[] args) throws IOException {
        read();
    }
}
