package com.asiainfo.excel.core;

import com.asiainfo.excel.inter.ExcelDesc;
import com.asiainfo.excel.model.SortableField;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

//通过类获取类所有的属性，根据@ExcelDesc注解中的order进行排序
public class GenObjectFields {

    public static List<SortableField> init(Class aClass)throws IntrospectionException{
        List<SortableField> sortableFields = new ArrayList<>();
        //获取所有的属性
        Field[] fields = aClass.getDeclaredFields();
        for (Field field:fields){
            //获取类属性的注解
            ExcelDesc excelDesc = field.getAnnotation(ExcelDesc.class);
            if (excelDesc != null && !excelDesc.ignoreField()){
                //根据类属性获取属性对应的get方法，用户通过反射机制获取元素的值
                PropertyDescriptor pd = new PropertyDescriptor(field.getName(),
                        aClass);
                Method method = pd.getReadMethod();
                SortableField sortableField = new SortableField(excelDesc, field, method);
                sortableFields.add(sortableField);
            }
        }
        //对获取到的属性进行排序
        Collections.sort(sortableFields, new Comparator<SortableField>() {
            public int compare(SortableField o1, SortableField o2) {
                return o1.getExcelDesc().order()-o2.getExcelDesc().order();
            }
        });
        return sortableFields;
    }
}
