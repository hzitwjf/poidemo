import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by wjf on 2016/11/08.
 * poi这个包是阿帕奇旗下一个用Java操作excel表格的包
 */
public class Demo3 {
    /**
     * 主要学习HSSF
     * @param args
     */
    public static void main(String[] args) {
        try {
            /*
             * Workbook是一个接口
             * HSSFWorkbook是一个实现类
             * 所以我们new出了他的实现类
             */
            Workbook workbook=new HSSFWorkbook();
            //使用文件输出流输出文件
            FileOutputStream fileOutputStream=new FileOutputStream("D:/迅雷下载/用poi制作的工作薄");
            //workbook.write方法：把内容写到流里面去
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
