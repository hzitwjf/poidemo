package com.wjfandmx.seconddemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * Created by wjf on 2016/11/8.
 * 创建一个sheet页
 */
public class Demo4 {
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
             * 值是一个String类型的数据
             */
            cell.setCellValue("时间单元格");
            //给第一列第二个单元格附上初始值new Date()看看结果如何；
            row.createCell(1).setCellValue(new Date());
            /*
             *给第一列第三个单元格赋值，值是一个用户可以接受的时间格式
             * 以下是做法一。
             */
            CellStyle cellStyle=workbook.createCellStyle();//创建一个单元格样式类,CellStyle
            DataFormat dataFormat=workbook.createDataFormat();//获取一个强转时间的类，DateFormat
            short time=dataFormat.getFormat("yyyy-mm-dd hh:mm:ss");//设置强转过后的时间格式
            cellStyle.setDataFormat(time);//把定义好的格式写入单元格样式类中
            cell=row.createCell(2);
            cell.setCellValue(new Date());
            cell.setCellStyle(cellStyle);
            /*
             *给第一列第四个单元格赋值，值是一个用户可以接受的时间格式
             * 以下是做法二。
             */
            CreationHelper creationHelper=workbook.getCreationHelper();//创建一个新的数据华格式
            cellStyle=workbook.createCellStyle();//创建一个单元格样式类,CellStyle
            dataFormat=creationHelper.createDataFormat();//获取一个强转时间的类，DateFormat
            time=dataFormat.getFormat("yyyy-mm-dd hh:mm:ss");//设置强转过后的时间格式
            cellStyle.setDataFormat(time);
            cell=row.createCell(3);
            cell.setCellValue(new Date());
            cell.setCellStyle(cellStyle);
            //给第一列第五个单元格赋值，值是一个用户可以接受的时间格式，做法三：
            cell=row.createCell(4);
            cell.setCellValue(Calendar.getInstance());//从日历里面获取时间
            cell.setCellStyle(cellStyle);

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
