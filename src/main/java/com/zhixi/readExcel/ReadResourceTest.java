package com.zhixi.readExcel;

import org.junit.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author zhangzhixi
 * @version 1.0
 * @date 2021-12-28 15:57
 */
public class ReadResourceTest {

    @Test
    public void testResourceRead() throws IOException {

        // 第一种：使用Spring提供的ClassPathResource，有没有斜杆都可以（推荐，功能都封装好了）
        ClassPathResource classPathResource = new ClassPathResource("static/excel/student_info.xlsx");
        InputStream inputStream = classPathResource.getInputStream();
        System.out.println(inputStream);

        // 第二种：使用Class#getResourceAsStream()，要加/
        InputStream classResource = this.getClass().getResourceAsStream("/static/excel/student_info.xlsx");
        System.out.println(classResource);

        // 第三种：使用ClassLoader#getResourceAsStream()，不加/
        InputStream classLoaderResource = this.getClass().getClassLoader().getResourceAsStream("static/excel/student_info.xlsx");
        System.out.println(classLoaderResource);
    }
}