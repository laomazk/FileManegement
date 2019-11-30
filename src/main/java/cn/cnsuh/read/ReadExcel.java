package cn.cnsuh.read;

import cn.cnsuh.domain.Student;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.metadata.Sheet;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Magic
 * @create 2019-11-30 14:20
 */

public class ReadExcel {
    public void run(int no) throws Exception {
        InputStream in = null;

        try {

            in = new FileInputStream(new File("d:/work/成绩表.xlsx"));

            Listener listener = new Listener();
            ExcelReader reader = new ExcelReader(in,null,listener);

            //读取excel
            Sheet sheet = new Sheet(no,0, Student.class);
            reader.read(sheet);


        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if(in!=null){
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

//    public static void main(String[] args) throws Exception {
//        new ReadExcel().run();
//    }
}
