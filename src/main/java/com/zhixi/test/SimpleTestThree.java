package com.zhixi.test;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author zhangzhixi
 * @version 1.0
 * @date 2021-12-28 14:50
 */
public class SimpleTestThree {

    @Test
    public void testSimpleWrite() throws IOException, ParseException {
        // 创建工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建工作表
        XSSFSheet sheet = workbook.createSheet("student");
        // 构造假数据
        List<Student> studentList = new ArrayList<>();
        studentList.add(new Student(1L, "周星驰", 58, "香港", new SimpleDateFormat("yyyy-MM-dd").parse("1962-6-22"), 174.0, false));
        studentList.add(new Student(2L, "李健", 46, "哈尔滨", new SimpleDateFormat("yyyy-MM-dd").parse("1974-9-23"), 174.5, true));
        studentList.add(new Student(3L, "周深", 28, "贵州", new SimpleDateFormat("yyyy-MM-dd").parse("1992-9-29"), 161.0, true));

        for (int i = 0; i < studentList.size(); i++) {
            Student student = studentList.get(i);
            // 创建行
            XSSFRow row = sheet.createRow(i);
            // 在当前行创建6个单元格，并设置数据（id不导出）
            row.createCell(0).setCellValue(student.getName());
            row.createCell(1).setCellValue(student.getAge());
            row.createCell(2).setCellValue(student.getAddress());
            row.createCell(3).setCellValue(student.getBirthday());
            row.createCell(4).setCellValue(student.getHeight());
            row.createCell(5).setCellValue(student.getIsMainlandChina());
        }

        FileOutputStream out = new FileOutputStream("H:\\→桌面←\\student_info_export.xlsx");
        workbook.write(out);
        out.flush();
        out.close();
        workbook.close();
        System.out.println("导出成功！");
    }

    @Data
    @AllArgsConstructor
    static class Student {
        private Long id;
        private String name;
        private Integer age;
        private String address;
        private Date birthday;
        private Double height;
        private Boolean isMainlandChina;
    }

}
