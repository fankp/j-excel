package com.asiainfo.excel.inter;

import java.lang.annotation.*;

// 注解会在class字节码文件中存在，在运行时可以通过反射获取到
@Retention(RetentionPolicy.RUNTIME)
//定义注解的作用目标**作用范围字段、枚举的常量/方法
@Target({ElementType.FIELD,ElementType.METHOD})
//说明该注解将被包含在javadoc中
@Documented
public @interface ExcelDesc {

    //是否被忽略
    boolean ignoreField() default false;
    //列的标题名称
    String cellName();
    //列的宽度
    int cellWidth() default 6000;
    //在excel中的顺序
    int order();
    //元素为空时写入excel的值
    String defaultNullVal() default "";
}
