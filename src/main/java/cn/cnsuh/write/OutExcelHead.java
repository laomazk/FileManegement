package cn.cnsuh.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * @author Magic
 * @create 2019-11-29 22:08
 */
public class OutExcelHead extends BaseRowModel {
    //定义属性,指定列的关系

    /**
     * value:指定表头内容
     * index:指定列的索引, 从 0开始
     */
    @ExcelProperty(value = "学号",index = 0)
    private String number;

    @ExcelProperty(value = "姓名",index = 1)
    private String name;

    @ExcelProperty(value = "成绩",index = 2)
    private String grade;



    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGrade() {
        return grade;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }
}
