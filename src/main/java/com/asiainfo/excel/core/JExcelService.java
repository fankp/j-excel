package com.asiainfo.excel.core;

import com.alibaba.fastjson.JSON;
import com.asiainfo.excel.model.SortableField;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

public class JExcelService {

    private String sheetName;
    //要写入excel表格的数据，数据格式为JSON
    private String excelData;
    //要生成excel的对象的类
    private Class aClass;

    public JExcelService(String excelData, Class aClass) {
        this.excelData = excelData;
        this.aClass = aClass;
        this.sheetName = "sheet1";
    }

    public JExcelService(String sheetName, String excelData, Class aClass) {
        this.sheetName = sheetName;
        this.excelData = excelData;
        this.aClass = aClass;
    }

    /**
     * 生成指定文件名的excel
     * @param fileName 要生成的excel文件名称
     * @throws IntrospectionException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws IOException
     */
    public void createExcel(String fileName) throws IntrospectionException, InvocationTargetException,
            IllegalAccessException, IOException {
        if (null == fileName || fileName.length() == 0){
            throw new RuntimeException("file name can't be null or empty.");
        }
        FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
        createExcel(fileOutputStream);
        fileOutputStream.close();
    }

    /**
     * 把创建的excel写入指定的输出流
     * @param out 生成的excel要输出的目标流
     * @throws IntrospectionException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws IOException
     */
    public void createExcel(OutputStream out) throws IntrospectionException, InvocationTargetException,
            IllegalAccessException, IOException {
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
        for (int i=0; i<sortableFields.size(); i++){
            String cellVal = sortableFields.get(i).getExcelDesc().cellName();
            HSSFCell cell = headRow.createCell(i);
            cell.setCellType(CellType.STRING);
            sheet.setColumnWidth(i, sortableFields.get(i).getExcelDesc().cellWidth());
            //填写数据
            cell.setCellStyle(headStyle);
            cell.setCellValue(cellVal);
        }
        for (int i=0; i<excelDataList.size(); i++){
            HSSFRow row = sheet.createRow(i + 1);
            for (int j=0; j<sortableFields.size(); j++){
                Cell cell = row.createCell(j);
                Method method = sortableFields.get(j).getReadMethod();
                Object cellVal =  method.invoke(excelDataList.get(i));
                cell.setCellValue(null == cellVal ? sortableFields.get(j).getExcelDesc().defaultNullVal():cellVal.toString());
            }
        }
        workbook.write(out);
        workbook.close();
    }
}
