package com.wjfandmx.firstdemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by wjf on 2016/11/8.
 * 创建一个sheet页
 */
public class Demo2 {
    public static void main(String[] args) {
        try {
            //定义一个新的工作簿
            Workbook workbook=new HSSFWorkbook();
            //创建一个sheet页
            workbook.createSheet("第一个sheet页");
            workbook.createSheet("第二个sheet页");
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
