package cn.cnsuh;

import cn.cnsuh.read.CopyExcel;
import cn.cnsuh.read.ReadExcel;
import cn.cnsuh.read.UpdateExcel;

import java.util.Scanner;

import static cn.cnsuh.write.OutExcel.createExcel;

/**
 * @author Magic
 * @create 2019-11-29 23:03
 */
public class main {
    public static void main(String[] args) throws Exception {
        Scanner sc = new Scanner(System.in);

        while (true) {
            System.out.println("学生成绩管理");
            System.out.println("1.将成绩输入到文件---2.查询文件中的成绩---3.修改指定成绩---4.将一个班的成绩复制到年级成绩表---5.退出");
            int n = sc.nextInt();
            switch (n) {
                case 1://写文件
                    createExcel();
                    break;
                case 2://查询
                    System.out.println("你要查询哪张成绩表，输入（1,2,3...）");
                    new ReadExcel().run(sc.nextInt());
                    break;
                case 3://修改
                    System.out.println("选择班级（输入1班成绩表,2班成绩表,3班成绩表...）");
                    String sheetname = sc.next();
                    System.out.println("输入学号");
                    String sid = sc.next();
                    System.out.println("输入新成绩");
                    int newgrade = sc.nextInt();
                    UpdateExcel.update(sheetname, sid, newgrade);
                    break;
                case 4://复制
                    System.out.println("选择复制的班级（输入1班成绩表,2班成绩表,3班成绩表...）");
                    String sheetname1 = sc.next();
                    System.out.println("选择插入表格");
                    String copyFile = sc.next();
                    CopyExcel.copy(sheetname1,copyFile);
                    break;
                case 5://退出
                    return;
                default:
                    System.out.println("输入错误,程序结束");
                    return;
            }


        }
    }
}
