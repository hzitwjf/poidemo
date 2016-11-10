package com.wjfandmx.firstdemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by wjf on 2016/11/8.
 * 创建一个sheet页
 */
public class Demo3 {
    public static void main(String[] args) {
        try {
            //定义一个新的工作簿
            Workbook workbook=new HSSFWorkbook();
            /*
             *创建一个sheet页
             * 返回一个Sheet对象，我们要去接收他
             */
            Sheet sheet=workbook.createSheet("第一个sheet页");
            /*
             *创建行，参数是下标，从第0行开始
             * 返回值类型是Row，需要接收
             */
            Row row=sheet.createRow(0);
            /*
             * 创建列，行加上列就是excel表格里面的单元格
             * 返回值类型是Cell，需要接收
             * 0代表第一列
             */
            Cell cell=row.createCell(0);
            /*
             * 给第一列第一个单元格赋值
             * 值是林黛玉
             */
            cell.setCellValue("林黛玉");
            //给第一列第二个单元格赋值，值是贾宝玉
            row.createCell(1).setCellValue("贾宝玉");
            //给第一列第三个单元格赋值，值是贾宝玉
            row.createCell(2).setCellValue("薛宝钗");
            //给第一列第四个单元格赋值，值是贾宝玉
            row.createCell(3).setCellValue("红楼梦");
            //给第一列第五个单元格赋值，值是一个数字类型
            row.createCell(4).setCellValue(5);
            //给第一列第六个单元格赋值，值是一个布尔类型
            row.createCell(5).setCellValue(true);
            //使用文件输出流输出文件
            FileOutputStream fileOutputStream = new FileOutputStream("D:/迅雷下载/用poi制作的工作薄.xls");
            //workbook.write方法：把内容写到流里面去
            workbook.write(fileOutputStream);
            //关闭文件输出流
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
