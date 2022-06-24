package org.libj.excel;

import com.toddfast.util.convert.TypeConverter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.libj.excel.anno.EnableExport;
import org.libj.excel.anno.EnableExportField;
import org.libj.excel.anno.EnableSelectList;
import org.libj.excel.anno.ImportIndex;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 详细使用说明请参考 App.java和Student.java
 */

public class ExcelHandler {
    private ExcelHandler() {
    }

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelHandler.class);
    private static final int MAX_ROW_NUMBER = 65530;

    /**
     * 将Excel转换为对象集合
     * 默认只读取第0个sheet
     *
     * @param excel Excel 文件
     * @param clazz pojo类型
     * @return 对象集合
     */
    public static <T> List<T> parseExcelToList(File excel, Class<T> clazz) throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        return parseExcelToList(excel, clazz, 0);
    }

    /**
     * 将Excel转换为对象集合
     *
     * @param excel    Excel 文件
     * @param clazz    pojo类型
     * @param sheetIdx sheet索引
     * @return 对象集合
     */
    public static <T> List<T> parseExcelToList(File excel, Class<T> clazz, int sheetIdx) throws IOException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        List<T> res = new ArrayList<>();
        try (InputStream is = Files.newInputStream(Paths.get(excel.getAbsolutePath())); Workbook workbook = WorkbookFactory.create(is)) {
            Sheet sheet = workbook.getSheetAt(sheetIdx);
            if (sheet != null) {
                int rownum = 2;
                Row row = sheet.getRow(rownum);
                while (row != null) {
                    String[] values = getValues(row);
                    T obj = getT(clazz, values);
                    res.add(obj);
                    row = sheet.getRow(++rownum);
                }
            }
        }
        return res;
    }

    /**
     * @param outputStream  输出流
     * @param dataList      导出数据
     * @param clazz         导出数据类型
     * @param selectListMap 转义字段
     */
    public static <T> void exportExcel(OutputStream outputStream, List<T> dataList, Class<T> clazz, Map<Integer, Map<String, String>> selectListMap) {
        if (!clazz.isAnnotationPresent(EnableExport.class))
            throw new RuntimeException(clazz + "can't be exported to excel file. consider add @EnableExport to the class");

        try (Workbook workbook = new SXSSFWorkbook()) {
            List<Field> fieldList = new ArrayList<>();
            List<String> colNames = new ArrayList<>();
            List<Integer> colWidths = new ArrayList<>();

            for (Field field1 : clazz.getDeclaredFields()) {
                if (field1.isAnnotationPresent(EnableExportField.class)) {
                    EnableExportField enableExportField = field1.getAnnotation(EnableExportField.class);
                    colNames.add(enableExportField.colName());
                    colWidths.add(enableExportField.colWidth());
                    fieldList.add(field1);
                }
            }
            int sheetCount = (dataList.size() / MAX_ROW_NUMBER) + 1;
            LOGGER.info("{} sheets will be created.", sheetCount);

            for (int sheetIdx = 0; sheetIdx < sheetCount; sheetIdx++) {
                LOGGER.info("exporting sheet {}.", sheetIdx);
                Sheet sheet = workbook.createSheet();
                sheet.setDefaultRowHeight((short) 400);
                createTitle(sheet, colNames.size() - 1, clazz.getAnnotation(EnableExport.class).filename());
                createHeadRow(sheet, colNames, colWidths);

                int rownum = 0;
                int fromIdx = sheetIdx * MAX_ROW_NUMBER;
                int toIdx = Math.min((sheetIdx + 1) * MAX_ROW_NUMBER, dataList.size());
                for (int dataIdx = fromIdx; dataIdx < toIdx; dataIdx++) {
                    Object obj = dataList.get(dataIdx);
                    Row row = sheet.createRow(rownum + 2);
                    exportRow(clazz, fieldList, selectListMap, row, obj);
                    rownum++;
                }

                createDataValidation(sheet, selectListMap);
            }

            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static <T> void exportRow(Class<T> clazz, List<Field> fieldList, Map<Integer, Map<String, String>> selectListMap, Row row, Object obj) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        for (int fieldIdx = 0; fieldIdx < fieldList.size(); fieldIdx++) {
            Field field = fieldList.get(fieldIdx);
            field.setAccessible(true);
            Object value = field.get(obj);
            EnableExportField enableExportField = field.getAnnotation(EnableExportField.class);
            String getMethodName = enableExportField.getter();
            if (!"".equals(getMethodName)) {
                Method method = clazz.getMethod(getMethodName);
                method.setAccessible(true);
                value = method.invoke(obj);
            }

            if (field.isAnnotationPresent(EnableSelectList.class)) {
                value = selectListMap.get(fieldIdx).get(value.toString());
            }
            row.createCell(fieldIdx).setCellValue(String.valueOf(value));
        }
    }

    /**
     * 读取行
     *
     * @param row 行
     * @return 读取行 封装为字符串数组
     */
    private static String[] getValues(Row row) {
        int cellNum = row.getPhysicalNumberOfCells();
        String[] values = new String[cellNum];
        for (int cellIdx = 0; cellIdx < cellNum; cellIdx++) {
            Cell cell = row.getCell(cellIdx);
            if (cell != null) {
                String value = cell.getStringCellValue();
                values[cellIdx] = value;
            }
        }
        return values;
    }

    /**
     * 将字符串数组封装到指定类型的字段上
     *
     * @param clazz  对象类型
     * @param values 用于封装的字符串数组
     * @param <T>    泛型类型
     * @return 封装好的对象
     * @throws IllegalAccessException    访问受限
     * @throws NoSuchMethodException     没有指定的setter方法
     * @throws InvocationTargetException 反射异常
     */
    private static <T> T getT(Class<T> clazz, String[] values) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        Field[] fields = clazz.getDeclaredFields();
        T obj;
        try {
            obj = clazz.newInstance();
        } catch (InstantiationException e) {
            throw new RuntimeException("该对象类型没有默认无参构造函数，请创建无参构造函数后重试！", e);
        }
        for (Field f : fields) {
            if (f.isAnnotationPresent(ImportIndex.class)) {
                ImportIndex annotation = f.getAnnotation(ImportIndex.class);
                int index = annotation.index();
                String setter = annotation.setter();
                Object val;
                try {
                    val = TypeConverter.convert(f.getType(), values[index]);
                } catch (Exception e) {
                    val = values[index];
                }
                if (!"".equals(setter)) {
                    Method method = clazz.getMethod(setter, f.getType());
                    method.setAccessible(true);
                    method.invoke(obj, val);
                } else {
                    f.setAccessible(true);
                    f.set(obj, val);
                }
            }
        }
        return obj;
    }

    /**
     * 创建一个跨列的标题行
     *
     * @param sheet     工作表
     * @param allColNum 合并的列数
     * @param title     标题
     */
    private static void createTitle(Sheet sheet, int allColNum, String title) {
        CellRangeAddress cra = new CellRangeAddress(0, 0, 0, allColNum);
        sheet.addMergedRegion(cra);
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(title);
    }

    /**
     * 设置表头标题栏以及表格高度
     *
     * @param sheet    工作表
     * @param colNames 表头列名
     */
    private static void createHeadRow(Sheet sheet, List<String> colNames, List<Integer> colWidths) {
        sheet.createFreezePane(0, 2);
        Row row = sheet.createRow(1);
        for (int i = 0; i < colNames.size(); i++) {
            Cell cell = row.createCell(i);
            sheet.setColumnWidth(i, colWidths.get(i) * 20);
            cell.setCellValue(colNames.get(i));
        }
    }

    /**
     * excel添加下拉数据校验
     *
     * @param sheet         待添加校验的sheet页
     * @param selectListMap 下拉框内容键值对
     */
    private static void createDataValidation(Sheet sheet, Map<Integer, Map<String, String>> selectListMap) {
        if (selectListMap != null) {
            for (Map.Entry<Integer, Map<String, String>> entry : selectListMap.entrySet()) {
                Integer key = entry.getKey();
                Map<String, String> value = entry.getValue();
                if (value.size() > 0) {
                    int i = 0;
                    String[] valueArr = new String[value.size()];
                    for (Map.Entry<String, String> ent : value.entrySet()) {
                        valueArr[i] = ent.getValue();
                        i++;
                    }
                    CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(2, 65535, key, key);
                    DataValidationHelper helper = sheet.getDataValidationHelper();
                    DataValidationConstraint constraint = helper.createExplicitListConstraint(valueArr);
                    DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);
                    if (dataValidation instanceof XSSFDataValidation) {
                        dataValidation.setSuppressDropDownArrow(true);
                        dataValidation.setShowErrorBox(true);
                    } else {
                        dataValidation.setSuppressDropDownArrow(false);
                    }
                    dataValidation.setEmptyCellAllowed(true);
                    dataValidation.setShowPromptBox(false);
                    dataValidation.createPromptBox("提示", "请选择下拉框中的数据");
                    sheet.addValidationData(dataValidation);
                }
            }
        }
    }
}
