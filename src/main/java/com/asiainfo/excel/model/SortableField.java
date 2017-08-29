package com.asiainfo.excel.model;

import com.asiainfo.excel.inter.ExcelDesc;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class SortableField {

    private ExcelDesc excelDesc;

    private Field field;

    private Method readMethod;

    public SortableField(ExcelDesc excelDesc) {
        this.excelDesc = excelDesc;
    }

    public SortableField(ExcelDesc excelDesc, Field field, Method readMethod) {
        this.excelDesc = excelDesc;
        this.field = field;
        this.readMethod = readMethod;
    }

    public ExcelDesc getExcelDesc() {
        return excelDesc;
    }

    public void setExcelDesc(ExcelDesc excelDesc) {
        this.excelDesc = excelDesc;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public Method getReadMethod() {
        return readMethod;
    }

    public void setReadMethod(Method readMethod) {
        this.readMethod = readMethod;
    }
}
