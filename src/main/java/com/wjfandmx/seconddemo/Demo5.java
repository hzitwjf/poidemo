package com.wjfandmx.seconddemo;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

/**
 * Created by wjf on 2016/11/8.
 * 创建一个sheet页
 */
public class Demo5 {
    public static void main(String[] args) {
        try {
            //创建文件输入流
            InputStream inputStream = new FileInputStream("D:/迅雷下载/红楼梦人物.xls");
            //这个类可以在inputStream里读取数据
            POIFSFileSystem poifsFileSystem=new POIFSFileSystem(inputStream);
            //创建一个新的工作簿，把读取到的数据放入新建的工作簿里面
            HSSFWorkbook workbook=new HSSFWorkbook(poifsFileSystem);
            //获取第一个sheet页
            HSSFSheet hssfSheet=workbook.getSheetAt(0);
            //判断一下，sheet页是否为空，如果为空，直接退出整个方法
            if (hssfSheet==null){
                return;
            }
            //如果Sheet页不为空，循环遍历所有行
            for (int rowNum=0;rowNum<=hssfSheet.getLastRowNum();rowNum++){
                //循环拿到每一行的值
                HSSFRow hssfRow=hssfSheet.getRow(rowNum);
                //判断当前行是否为空，如果为空，跳过本次循环；
                if (hssfRow==null){
                    continue;
                }
                //循环遍历所有的列
                for (int cellNum=0;cellNum<=hssfRow.getLastCellNum();cellNum++){
                    //循环拿到所有列的值
                    HSSFCell hssfCell=hssfRow.getCell(cellNum);
                    //判断当前列是否为空，如果为空，跳过本次循环；
                    if (hssfCell==null){
                        continue;
                    }
                    System.out.print("\t\t"+getValue(hssfCell));
                }
                System.out.println();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 因为列的值有多数字，布尔，字符串，日期类型
     * 所以声明一个方法专门用来处理数据类型
     * main方法无法调用非静态方法，所以方法是静态的
     * @param hssfCell  传入一个列
     * @return  该列的值
     */
    /*private static String getCellValue(HSSFCell hssfCell){
        DecimalFormat df = new DecimalFormat("#");
        if (hssfCell == null)
            return "";
        switch (hssfCell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                if(HSSFDateUtil.isCellDateFormatted(hssfCell)){
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                    return sdf.format(HSSFDateUtil.getJavaDate(hssfCell.getNumericCellValue())).toString();
                }
                return df.format(hssfCell.getNumericCellValue());
            case HSSFCell.CELL_TYPE_STRING:
                System.out.println(hssfCell.getStringCellValue());
                return hssfCell.getStringCellValue();
            case HSSFCell.CELL_TYPE_FORMULA:
                return hssfCell.getCellFormula();
            case HSSFCell.CELL_TYPE_BLANK:
                return "";
            case HSSFCell.CELL_TYPE_BOOLEAN:
                return String.valueOf(hssfCell.getBooleanCellValue());
            case HSSFCell.CELL_TYPE_ERROR:
                return String.valueOf(hssfCell.getErrorCellValue());
        }
        return "";
    }*/
    private static String getValue(HSSFCell hssfCell){
        if(hssfCell.getCellType()==HSSFCell.CELL_TYPE_BOOLEAN){
            return String.valueOf(hssfCell.getBooleanCellValue());
        }else if(hssfCell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
            if (HSSFDateUtil.isCellDateFormatted(hssfCell)){
                return String.valueOf(hssfCell.getDateCellValue());
            }else {
                return String.valueOf(hssfCell.getNumericCellValue());
            }
        }else if (hssfCell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
            return String.valueOf(hssfCell.getDateCellValue());
        }else {
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }
}
