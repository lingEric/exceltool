## 使用说明

通过给实体类标记注解，然后调用一行代码实现导入导出到表格

需要用到的注解有：

- `EnableExport` 标注了该注解的类才可以导出

- `EnableExportField` 标注了该注解的字段才会导出
- `EnableSelectList`是否需要设置为下拉框
- `ImportIndex ` 标注了该注解 才会导入到实体类的相应字段中



具体使用请看下面的演示。



## 实体类

```java
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
```



## 导入导出操作

```java
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
```

