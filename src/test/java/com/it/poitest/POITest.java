package com.it.poitest;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class POITest {

    @Test
    public void readTest01() throws Exception {
        //1.获取一个工作薄，并且把excel的输入流对象传递进来。  03Excel版本对应的实现类是HSSFWorkbook
        Workbook workbook = new HSSFWorkbook(new FileInputStream("D:\\黑马\\就业班\\就业班-day58(SaaS-Export第10天)\\day10_saas_export\\03.资料和工具\\01 excel资源模板\\上传货物模板2.xls"));

        //获取一个工作单
        Sheet sheet = workbook.getSheetAt(0);

        //获取行
        Row row = sheet.getRow(0);

        System.out.println(row.getCell(1).getStringCellValue()+"\t");
        System.out.println(row.getCell(2).getStringCellValue()+"\t");
        System.out.println(row.getCell(3).getStringCellValue()+"\t");
        System.out.println(row.getCell(4).getStringCellValue()+"\t");
        System.out.println(row.getCell(5).getStringCellValue()+"\t");
        System.out.println(row.getCell(6).getStringCellValue()+"\t");
        System.out.println(row.getCell(7).getStringCellValue()+"\t");
        System.out.println(row.getCell(8).getStringCellValue()+"\t");
        System.out.println(row.getCell(9).getStringCellValue());

    }

    @Test
    public void readTest2() throws Exception {
        //1.获取一个工作薄，并且把excel的输入流对象传递进来。  03Excel版本对应的实现类是HSSFWorkbook
        Workbook workbook = new HSSFWorkbook(new FileInputStream("D:\\黑马\\就业班\\就业班-day58(SaaS-Export第10天)\\day10_saas_export\\03.资料和工具\\01 excel资源模板\\上传货物模板2.xls"));

        //获取一个工作单
        Sheet sheet = workbook.getSheetAt(0);

        //获取行
        Row row = sheet.getRow(0);

        System.out.println(row.getCell(1).getStringCellValue()+"\t");
        System.out.println(row.getCell(2).getStringCellValue()+"\t");
        System.out.println(row.getCell(3).getStringCellValue()+"\t");
        System.out.println(row.getCell(4).getStringCellValue()+"\t");
        System.out.println(row.getCell(5).getStringCellValue()+"\t");
        System.out.println(row.getCell(6).getStringCellValue()+"\t");
        System.out.println(row.getCell(7).getStringCellValue()+"\t");
        System.out.println(row.getCell(8).getStringCellValue()+"\t");
        System.out.println(row.getCell(9).getStringCellValue());

        System.out.println("==================================================================");

        //获取excel表格的总行数
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i <rows ; i++) {
            row = sheet.getRow(i);

            System.out.println(row.getCell(1).getStringCellValue()+"\t");
            System.out.println(row.getCell(2).getStringCellValue()+"\t");
            System.out.println(row.getCell(3).getNumericCellValue()+"\t");
            System.out.println(row.getCell(4).getStringCellValue()+"\t");
            System.out.println(row.getCell(5).getNumericCellValue()+"\t");
            System.out.println(row.getCell(6).getNumericCellValue()+"\t");
            System.out.println(row.getCell(7).getNumericCellValue()+"\t");
            System.out.println(row.getCell(8).getStringCellValue()+"\t");
            System.out.println(row.getCell(9).getStringCellValue());
        }

    }

    /*
    读取07(xls)版本的多行数据
   */
    @Test
    public void readTest03() throws Exception {
        //1.获取一个工作薄，并且把excel的输入流对象传递进来。  03Excel版本对应的实现类是HSSFWorkbook
        Workbook workbook = new XSSFWorkbook(new FileInputStream("D:\\黑马\\就业班\\就业班-day58(SaaS-Export第10天)\\day10_saas_export\\03.资料和工具\\01 excel资源模板\\上传货物模板.xlsx"));
        //2.获取一个工作单
        Sheet sheet = workbook.getSheetAt(0);

        //获取行
        Row row = sheet.getRow(0);


        System.out.println(row.getCell(1).getStringCellValue()+"\t");
        System.out.println(row.getCell(2).getStringCellValue()+"\t");
        System.out.println(row.getCell(3).getStringCellValue()+"\t");
        System.out.println(row.getCell(4).getStringCellValue()+"\t");
        System.out.println(row.getCell(5).getStringCellValue()+"\t");
        System.out.println(row.getCell(6).getStringCellValue()+"\t");
        System.out.println(row.getCell(7).getStringCellValue()+"\t");
        System.out.println(row.getCell(8).getStringCellValue()+"\t");
        System.out.println(row.getCell(9).getStringCellValue());

        System.out.println("==================================================================");

        //获取excel表格的总行数
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i <rows ; i++) {
            row = sheet.getRow(i);

            System.out.println(row.getCell(1).getStringCellValue()+"\t");
            System.out.println(row.getCell(2).getStringCellValue()+"\t");
            System.out.println(row.getCell(3).getNumericCellValue()+"\t");
            System.out.println(row.getCell(4).getStringCellValue()+"\t");
            System.out.println(row.getCell(5).getNumericCellValue()+"\t");
            System.out.println(row.getCell(6).getNumericCellValue()+"\t");
            System.out.println(row.getCell(7).getNumericCellValue()+"\t");
            System.out.println(row.getCell(8).getStringCellValue()+"\t");
            System.out.println(row.getCell(9).getStringCellValue());
        }
    }

    /*
      生成excel表格
     */
    @Test
    public void writeTest() throws Exception {
        //1.创建一个工作薄
        Workbook workbook = new XSSFWorkbook();

        //2.创建一个工作单
        Sheet sheet = workbook.createSheet("小喇叭");

        //创建行
        Row row = sheet.createRow(0);

        //创建单元格
        row.createCell(1).setCellValue("姓名");
        row.createCell(2).setCellValue("年龄");
        row.createCell(3).setCellValue("性别");
        row.createCell(4).setCellValue("身高");

        workbook.write(new FileOutputStream("D:\\work\\work.xlsx"));
    }
}
