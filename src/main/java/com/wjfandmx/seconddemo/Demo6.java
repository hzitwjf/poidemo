package com.wjfandmx.seconddemo;


import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created by wjf on 2016/11/8.
 * 创建一个sheet页
 */
public class Demo6 {
    public static void main(String[] args) {
        try {
            //创建文件输入流
            InputStream inputStream = new FileInputStream("D:/迅雷下载/红楼梦人物.xls");
            //这个类可以在inputStream里读取数据
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
            //创建一个新的工作簿，把读取到的数据放入新建的工作簿里面
            HSSFWorkbook workbook = new HSSFWorkbook(poifsFileSystem);
            //专门用来提取的类，参数放工作簿
            ExcelExtractor excelExtractor=new ExcelExtractor(workbook);

            //getText方法指定抽取该excel表格中的文本
            System.out.println(excelExtractor.getText());
            /*
             * 如果不需要显示结果又sheet页的名字可设置
             * excelExtractor.setIncludeSheetNames(false);
             * 参数为Boolean类型，false，不需要显示sheet页名字，默认为true
             */
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
