package com.asiainfo.excel.core;

import com.alibaba.fastjson.JSON;
import com.asiainfo.excel.inter.ExcelDesc;
import com.asiainfo.excel.model.SortableField;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import sun.font.TrueTypeFont;

import java.beans.IntrospectionException;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.util.*;

public class JExcelService<E> {

    private String sheetName;
    //要写入excel表格的数据，数据格式为JSON
    private String excelData;
    //要生成excel的对象的类
    private Class aClass;

    public JExcelService(Class<E> aClass) {
        this.aClass = aClass;
        this.sheetName = "sheet1";
    }

    public JExcelService(String sheetName, Class<E> aClass) {
        this.sheetName = sheetName;
        this.aClass = aClass;
    }

    /**
     * 生成指定文件名的excel
     *
     * @param fileName 要生成的excel文件名称
     * @throws IntrospectionException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws IOException
     */
    public void createExcel(String fileName, String excelData) throws IntrospectionException, InvocationTargetException,
            IllegalAccessException, IOException {
        if (null == fileName || fileName.length() == 0) {
            throw new RuntimeException("file name can't be null or empty.");
        }
        FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
        createExcel(fileOutputStream, excelData);
        fileOutputStream.close();
    }

    /**
     * 把创建的excel写入指定的输出流
     *
     * @param out 生成的excel要输出的目标流
     * @throws IntrospectionException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws IOException
     */
    public void createExcel(OutputStream out, String excelData) throws IntrospectionException, InvocationTargetException,
            IllegalAccessException, IOException {
        this.excelData = excelData;
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(this.sheetName);
        //获取类的属性
        List<SortableField> sortableFields = GenObjectFields.init(aClass);
        //按照JSON的形式解析要写入excel的数据
        List excelDataList = JSON.parseArray(excelData, aClass);
        //标题行
        HSSFRow headRow = sheet.createRow(0);
        HSSFCellStyle headStyle = workbook.createCellStyle();
        headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());// 设置背景色
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置背景色
        HSSFFont headFont = workbook.createFont();
        headFont.setBold(true);
        headStyle.setAlignment(HorizontalAlignment.CENTER);
        headStyle.setFont(headFont);
        for (int i = 0; i < sortableFields.size(); i++) {
            String cellVal = sortableFields.get(i).getExcelDesc().cellName();
            HSSFCell cell = headRow.createCell(i);
            cell.setCellType(CellType.STRING);
            sheet.setColumnWidth(i, sortableFields.get(i).getExcelDesc().cellWidth());
            //填写数据
            cell.setCellStyle(headStyle);
            cell.setCellValue(cellVal);
        }
        for (int i = 0; i < excelDataList.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            for (int j = 0; j < sortableFields.size(); j++) {
                Cell cell = row.createCell(j);
                Method method = sortableFields.get(j).getReadMethod();
                Object cellVal = method.invoke(excelDataList.get(i));
                cell.setCellValue(null == cellVal ? sortableFields.get(j).getExcelDesc().defaultNullVal() : cellVal.toString());
            }
        }
        workbook.write(out);
        workbook.close();
    }
    public List<?> readExcel(String fileName) throws IntrospectionException, InvocationTargetException,
            IllegalAccessException, IOException, InstantiationException {
        if (null == fileName || fileName.length() == 0) {
            throw new RuntimeException("file name can't be null or empty.");
        }
        FileInputStream fileInputStream = new FileInputStream(new File(fileName));
        return readExcel(fileInputStream);

    }

    public List<?> readExcel(InputStream is) throws IntrospectionException, InvocationTargetException,
            IllegalAccessException, IOException, InstantiationException {
        //获取类的属性
        List<SortableField> sortableFields = GenObjectFields.init(aClass);
        Map<String, SortableField> textToKey = new HashMap<String, SortableField>();
        ExcelDesc _excel = null;
        for (SortableField field : sortableFields) {
            textToKey.put(field.getExcelDesc().cellName(), field);
        }
        HSSFWorkbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheet(this.sheetName);
        Row title = sheet.getRow(0);
        // 标题数组，后面用到，根据索引去标题名称，通过标题名称去字段名称用到 textToKey
        String[] titles = new String[title.getPhysicalNumberOfCells()];
        for (int i = 0; i < title.getPhysicalNumberOfCells(); i++) {
            titles[i] = title.getCell(i).getStringCellValue();
        }

        List<E> list = new ArrayList<>();

        int rowIndex = 0;
        int columnCount = titles.length;
        Cell cell = null;
        Row row = null;
        for (Iterator<Row> it = sheet.rowIterator(); it.hasNext(); ) {
            row = it.next();
            if (rowIndex++ == 0) {
                continue;
            }
            if (row == null) {
                break;
            }
            E e = (E) aClass.newInstance();

            for (int i = 0; i < columnCount; i++) {
                cell = row.getCell(i);
                String cellValue = null;
                switch(cell.getCellTypeEnum()){
                    case _NONE:
                        cellValue = null;
                        break;
                    case BLANK:
                        cellValue = null;
                        break;
                    case STRING:
                        cellValue = cell.getStringCellValue();
                        break;
                    case BOOLEAN:
                        cellValue = true == cell.getBooleanCellValue() ? "Y" : "N";
                        break;
                    case FORMULA:
                        cellValue = cell.getCellFormula();
                        break;
                    case NUMERIC:
                        DecimalFormat format = new DecimalFormat("###################.###########");
                        cellValue = format.format(cell.getNumericCellValue());
                    case ERROR:
                        break;

                }
                SortableField field = textToKey.get(titles[i]);
                field.getField().setAccessible(true);
                field.getField().set(e, cellValue);
            }
            list.add(e);
        }
        return list;
    }
}
