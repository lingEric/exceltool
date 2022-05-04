package org.labj.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * 演示如何使用导出导入excel
 * @author eric ling
 */
public class App {
    public static void main(String[] args) throws IOException, InvalidFormatException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        List<Student> stuList = new ArrayList<>();
        Random random = new Random();
        for (int i = 1; i < 655390; i++) {
            stuList.add(new Student(i + "号", random.nextInt(20) + 1, (i & 2) == 0 ? "男" : "女", new int[]{90, 90, 90}));
        }

        String filename = "某某中学全体学生名单.xls";
        ExcelHandler.exportExcel(Files.newOutputStream(Paths.get(filename)), stuList, Student.class, null);
        System.out.println(ExcelHandler.parseExcelToList(new File(filename), Student.class));
    }
}
