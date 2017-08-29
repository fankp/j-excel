# j-excel
## 1.功能介绍
        本工具的主要功能是把对象列表中的属性值输出到excel中。
        需要使用注解的形式在Java Bean中对excel的列头，列宽等基本信息进行描述。
## 2.实现原理
        1）通过Java注解把Excel中的列头信息绑定在Java Bean的属性上；
        2）通过Java类反射技术获取提供的类型中的属性、属性的get方法、属性对应的注解；
        3）通过获取到的表头信息，以及属性的get方法，获取清单中的单个元素，并把结果写入excel。
## 3.项目依赖
        1）alibaba json解析工具；
        2）apache poi操作excel工具；
