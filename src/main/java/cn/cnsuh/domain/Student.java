package cn.cnsuh.domain;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * @author Magic
 * @create 2019-11-30 14:15
 */
public class Student extends BaseRowModel {

    @ExcelProperty(index = 0)
    private String number;

    @ExcelProperty(index = 1)
    private String name;

    @ExcelProperty(index = 2)
    private String grade;


    public Student() {
    }

    public Student(String number, String name, String grade) {
        this.number = number;
        this.name = name;
        this.grade = grade;
    }

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

    @Override
    public String toString() {
        return "Student{" +
                "number='" + number + '\'' +
                ", name='" + name + '\'' +
                ", grade='" + grade + '\'' +
                '}';
    }
}
