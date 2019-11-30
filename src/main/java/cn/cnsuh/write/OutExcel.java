package cn.cnsuh.write;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * @author Magic
 * @create 2019-11-29 22:07
 */
public class OutExcel {
    static Scanner sc = new Scanner(System.in);
    public static int count;
    public static void createExcel() {
        OutputStream out = null;
        try {
            out = new FileOutputStream("d:/work/成绩表.xlsx");

//            Sheet sheet = new Sheet(1, 0, OutExcelHead.class);
//            sheet.setSheetName("一班成绩表");
//
//            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
//            writer.write0(getData(), sheet);
//            writer.finish();
//
//            System.out.println("生成excel完毕");
            System.out.println("请输入班级的数量，即创建sheet的个数");
            count = sc.nextInt();//创建的sheet数量
            Sheet[] sheets = new Sheet[count];
            ExcelWriter writer = new ExcelWriter(out,ExcelTypeEnum.XLSX);
            for (int i = 0; i < count ; i++) {
                sheets[i] = new Sheet(i+1,0,OutExcelHead.class);
                sheets[i].setSheetName((i+1)+"班成绩表");

                writer.write0(getData(),sheets[i]);
            }
            writer.finish();
            System.out.println("生成excel完毕");
            return;


        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    //获取输出到excel的数据List<List<String>>
    public static List<List<String>> getData() {
        List<List<String>> data = new ArrayList<List<String>>();
//        //第一行数据
//        List<String> row = new ArrayList<String>();
//        row.add("9000");
//        row.add("1000");
//        row.add("6080");
//        data.add(row);
//        System.out.println(count);

            System.out.println("请输入当前要输入的班级人数");
            int totalNum = sc.nextInt();
            for (int i = 0; i < totalNum; i++) {
                List<String> row = new ArrayList<String>();
                System.out.println("请输入第"+(i+1)+"个同学的学号");
                row.add(sc.next());
                System.out.println("请输入第"+(i+1)+"个同学的姓名");
                row.add(sc.next());
                System.out.println("请输入第"+(i+1)+"个同学的成绩");
                row.add(sc.next());
                data.add(row);
            }



        return data;

    }
}
