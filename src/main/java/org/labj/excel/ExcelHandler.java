package org.labj.excel;

import com.alibaba.fastjson.util.TypeUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.labj.excel.anno.EnableExport;
import org.labj.excel.anno.EnableExportField;
import org.labj.excel.anno.EnableSelectList;
import org.labj.excel.anno.ImportIndex;

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
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 详细使用说明请参考 App.java和Student.java
 */

public class ExcelHandler {
    private static final Logger LOGGER = Logger.getLogger(ExcelHandler.class.getName());
    private static final int MAX_ROW_NUMBER = 65530;
    private static HSSFCellStyle basicCellStyle;
    /**
     * 多次导出时，使用的basicCellStyle始终都是第一个工作簿的，setCellStyle()会抛出异常
     * 所以添加一个标记位，每次调用exportExcel()方法时，重置为true，然后重新创建basicCellStyle
     */
    private static boolean anotherOne = false;

    /**
     * 将Excel转换为对象集合
     * 默认只读取第0个sheet
     *
     * @param excel Excel 文件
     * @param clazz pojo类型
     * @return 对象集合
     */
    public static <T> List<T> parseExcelToList(File excel, Class<T> clazz) throws IOException, InvalidFormatException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
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
    public static <T> List<T> parseExcelToList(File excel, Class<T> clazz, int sheetIdx) throws IOException, InvalidFormatException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        List<T> res = new ArrayList<>();
        try (InputStream is = Files.newInputStream(Paths.get(excel.getAbsolutePath()))) {
            Workbook workbook = WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(sheetIdx);
            if (sheet != null) {
                int rownum = 2;  // 从第2行开始读取（第 0 行为标题，第 1 行为表头）
                Row row = sheet.getRow(rownum);
                while (row != null) {
                    String[] values = getValues(row);  // 读取行
                    T obj = getT(clazz, values);
                    res.add(obj);  // 封装对象属性
                    row = sheet.getRow(++rownum); // 读取下一行
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
        anotherOne = true;
        if (!clazz.isAnnotationPresent(EnableExport.class))
            throw new RuntimeException(clazz + "can't be exported to excel file.");

        HSSFWorkbook workbook = new HSSFWorkbook();  // 工作簿
        List<String> colNames = new ArrayList<>();  // 列名
        List<Integer> colWidths = new ArrayList<>(); // 列宽
        List<Field> fieldList = new ArrayList<>();  // 导出字段

        // 填充导出字段信息
        for (Field field1 : clazz.getDeclaredFields()) {  // 获取导出字段信息
            if (field1.isAnnotationPresent(EnableExportField.class)) {
                EnableExportField enableExportField1 = field1.getAnnotation(EnableExportField.class);
                colNames.add(enableExportField1.colName());
                colWidths.add(enableExportField1.colWidth());
                fieldList.add(field1);
            }
        }
        LOGGER.info("Fields need to be exported:" + fieldList);

        EnableExport export = clazz.getAnnotation(EnableExport.class);
        String titleCellValue = export.filename();
        int sheetCount = (dataList.size() / MAX_ROW_NUMBER) + 1;  // 计算sheet页数量
        LOGGER.info(sheetCount + " sheets will be created.");

        // 开始导出
        for (int sheetIdx = 0; sheetIdx < sheetCount; sheetIdx++) {  // 分页导出
            LOGGER.info("exporting sheet " + sheetIdx);
            HSSFSheet hssfsheet = workbook.createSheet();
            hssfsheet.setDefaultRowHeight((short) 400);
            createTitle(workbook, hssfsheet, colNames.size() - 1, titleCellValue);  // 表格第一行标题
            createHeadRow(workbook, hssfsheet, colNames, colWidths);  // 表格第二行表头

            try {
                int rownum = 0;  // 导出行号
                int fromIdx = sheetIdx * MAX_ROW_NUMBER;
                int toIdx = Math.min((sheetIdx + 1) * MAX_ROW_NUMBER, dataList.size());// 既要保证不超出页面最大值 也要保证不超出dataList索引范围
                for (int dataIdx = fromIdx; dataIdx < toIdx; dataIdx++) {
                    Object obj = dataList.get(dataIdx);
                    HSSFRow hssfRow = hssfsheet.createRow(rownum + 2);
                    exportRow(clazz, fieldList, selectListMap, getBasicCellStyle(workbook), hssfRow, obj);  // 导出一行
                    rownum++;
                }
                //创建下拉列表（该列仅可从下拉列表中选取数据输入）
                createDataValidation(hssfsheet, selectListMap);

            } catch (IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
                throw new RuntimeException("数据导出失败！！！", e);
            }
        }

        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static <T> void exportRow(Class<T> clazz, List<Field> fieldList, Map<Integer, Map<String, String>> selectListMap, HSSFCellStyle cellStyle, HSSFRow hssfRow, Object obj) throws IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        for (int fieldIdx = 0; fieldIdx < fieldList.size(); fieldIdx++) {
            Field field = fieldList.get(fieldIdx);
            field.setAccessible(true);
            Object value = field.get(obj);
            // field.getType().isArray();
            EnableExportField enableExportField = field.getAnnotation(EnableExportField.class);
            String getMethodName = enableExportField.getter();
            // getter覆盖字段值 可以在getter里面做一层抽象
            if (!"".equals(getMethodName)) {
                Method method = clazz.getMethod(getMethodName);
                method.setAccessible(true);
                value = method.invoke(obj);
            }

            // 字段转义
            if (field.isAnnotationPresent(EnableSelectList.class)) {
                if (selectListMap != null && selectListMap.get(fieldIdx) != null) {
                    String mapToValue = selectListMap.get(fieldIdx).get(value.toString());
                    if (mapToValue != null) value = mapToValue;
                }
            }
            setCellValue(value, hssfRow, cellStyle, fieldIdx);
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
        for (int cellIdx = 0; cellIdx <= cellNum; cellIdx++) {
            Cell cell = row.getCell(cellIdx);
            if (cell != null) {
                cell.setCellType(Cell.CELL_TYPE_STRING);
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
                    val = TypeUtils.cast(values[index], f.getType(), null);
                } catch (Exception e) {
                    val = values[index];
                }
                if (!"".equals(setter)) {  // 数组类型
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
     * @param workbook  工作簿
     * @param hssfsheet 工作表
     * @param allColNum 合并的列数
     * @param title     标题
     */
    private static void createTitle(HSSFWorkbook workbook, HSSFSheet hssfsheet, int allColNum, String title) {
        CellRangeAddress cra = new CellRangeAddress(0, 0, 0, allColNum);
        hssfsheet.addMergedRegion(cra);
        RegionUtil.setBorderBottom(2, cra, hssfsheet, workbook);
        RegionUtil.setBorderLeft(2, cra, hssfsheet, workbook);
        RegionUtil.setBorderRight(2, cra, hssfsheet, workbook);
        RegionUtil.setBorderTop(2, cra, hssfsheet, workbook);
        HSSFRow hssfRow = hssfsheet.createRow(0);
        HSSFCell hssfcell = hssfRow.createCell(0);
        hssfcell.setCellStyle(getBasicCellStyle(workbook));
        hssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
        hssfcell.setCellValue(title);
    }

    /**
     * 设置表头标题栏以及表格高度
     *
     * @param workbook  工作簿
     * @param hssfsheet 工作表
     * @param colNames  表头列名
     */
    private static void createHeadRow(HSSFWorkbook workbook, HSSFSheet hssfsheet, List<String> colNames, List<Integer> colWidths) {
        hssfsheet.createFreezePane(0, 2);
        HSSFRow hssfRow = hssfsheet.createRow(1);
        for (int i = 0; i < colNames.size(); i++) {
            HSSFCell hssfcell = hssfRow.createCell(i);
            hssfsheet.setColumnWidth(i, colWidths.get(i) * 20);
            hssfcell.setCellStyle(getBasicCellStyle(workbook));
            hssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
            hssfcell.setCellValue(colNames.get(i));
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
                // 第几列校验（0开始）key 数据源数组value
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
                    //处理Excel兼容性问题
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

    /**
     * 判断字符串是否为数字
     *
     * @param str 待验证的字符串
     * @return 是否为数字
     */
    private static boolean isNumeric(String str) {
        Pattern pattern = Pattern.compile("[-+]?((\\d+)([.](\\d+))?)$");
        if (str != null && !"".equals(str.trim())) {
            Matcher matcher = pattern.matcher(str);
            if (matcher.matches()) {
                return str.contains(".") || !str.startsWith("0");
            }
        }
        return false;
    }

    /**
     * 设置单元格的值
     *
     * @param value     单元格填充数据
     * @param hssfRow   行
     * @param cellStyle 单元格样式
     * @param cellIndex 单元格索引
     */
    private static void setCellValue(Object value, HSSFRow hssfRow, CellStyle cellStyle, int cellIndex) {
        String valueStr = String.valueOf(value);
        HSSFCell hssfcell = hssfRow.createCell(cellIndex);
        hssfcell.setCellStyle(cellStyle);
        if (isNumeric(valueStr)) {
            hssfcell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
            hssfcell.setCellValue(Double.parseDouble(valueStr));
        } else {
            hssfcell.setCellType(HSSFCell.CELL_TYPE_STRING);
            hssfcell.setCellValue(valueStr);
        }
    }

    /**
     * 单例模式获取单元格样式
     *
     * @param workbook 工作簿
     * @return 基本单元格样式
     */
    private static HSSFCellStyle getBasicCellStyle(HSSFWorkbook workbook) {
        if (basicCellStyle == null || anotherOne) {
            basicCellStyle = workbook.createCellStyle();
            basicCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            basicCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            basicCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
            basicCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
            basicCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            basicCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            basicCellStyle.setWrapText(true);
        }
        return basicCellStyle;
    }

}
