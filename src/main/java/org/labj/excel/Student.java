package org.labj.excel;

import org.labj.excel.anno.EnableExport;
import org.labj.excel.anno.EnableExportField;
import org.labj.excel.anno.ImportIndex;

import java.util.Arrays;
import java.util.stream.Collectors;

/**
 * 此对象类型用于演示如何使用
 *
 * @author eric ling
 */

@EnableExport(filename = "表格标题：某某中学全体学生名单")
public class Student {
    @ImportIndex(index = 0)  // 导入时列索引 从0开始
    @EnableExportField(colName = "姓名")
    private String name;

    @ImportIndex(index = 1)
    @EnableExportField(colName = "年龄")
    private int age;

    @ImportIndex(index = 2)
    @EnableExportField(colName = "性别")
    private String gender;

    @EnableExportField(colName = "语数外", colWidth = 800, getter = "getGradeString")  // 导出时指定getter方法，针对数组类型
    private int[] grade;

    @ImportIndex(index = 3)
    private String gradeString;

    /**
     * 导入表格时，对象类型必须要有无参构造函数
     */
    public Student() {
    }

    public Student(String name, int age, String gender, int[] grade) {
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.grade = grade;
    }

    /**
     * 导出时 调用此方法 获取字符串（如果不指定getter，则使用反射直接获取字段值）
     */
    public String getGradeString() {
        return Arrays.stream(grade).mapToObj(String::valueOf).collect(Collectors.joining(","));
    }

    @Override
    public String toString() {
        return "Student{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", gender='" + gender + '\'' +
                ", grade=" + Arrays.toString(grade) +
                ", gradeString='" + gradeString + '\'' +
                '}';
    }
}
