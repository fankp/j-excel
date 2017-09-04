package com.asiainfo.excel.test;

import com.alibaba.fastjson.JSON;
import com.asiainfo.excel.model.SortableField;
import com.asiainfo.excel.core.GenObjectFields;
import com.asiainfo.excel.core.JExcelService;
import com.asiainfo.excel.test.model.User;
import junit.framework.TestCase;

import java.beans.IntrospectionException;
import java.util.ArrayList;
import java.util.List;

public class TestGetFields extends TestCase {

    public void testGetClassFields() throws IntrospectionException{
        List<SortableField> sortableFields = GenObjectFields.init(User.class);
        for (SortableField sortableField:sortableFields){
            System.out.println("单元格标题：" + JSON.toJSONString(sortableField.getExcelDesc().cellName()));
            System.out.println("类属性名称：" + JSON.toJSONString(sortableField.getField().getName()));
            System.out.println("获取属性方法" + JSON.toJSONString(sortableField.getReadMethod().getName()));
        }
    }

    public void testGenExcelData() throws Exception{
        List<User> users = new ArrayList<>();
        User user = new User();
        user.setUserId("1");
        user.setUserName("fankp");
        user.setUserPwd("fankp");
        users.add(user);
        JExcelService jExcelService = new JExcelService("sheet2", User.class);
        jExcelService.createExcel("data/test.xls",JSON.toJSONString(users));
    }
    public void testReadExcelData() throws Exception{
        JExcelService jExcelService = new JExcelService<User>("sheet2", User.class);
        List<User> users = jExcelService.readExcel("data/test.xls");
        System.out.println(users);
    }
}
