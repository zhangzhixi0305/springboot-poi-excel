package com.zhixi.test;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;
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
 * @date 2021-12-28 15:08
 */
public class SimpleTestFive {

    /**
     * 按模板样式导出Excel
     *
     * @throws ParseException
     * @throws IOException
     */
    @Test
    public void testWriteWithStyle() throws IOException, ParseException {
        // 查询数据
        List<Student> studentList = new ArrayList<>();
        studentList.add(new Student(1L, "周深", 28, "贵州", new SimpleDateFormat("yyyy-MM-dd").parse("1992-9-29"), 161.0, true));
        studentList.add(new Student(2L, "李健", 46, "哈尔滨", new SimpleDateFormat("yyyy-MM-dd").parse("1974-9-23"), 174.5, true));
        studentList.add(new Student(3L, "周星驰", 58, "香港", new SimpleDateFormat("yyyy-MM-dd").parse("1962-6-22"), 174.0, false));

        // 读取模板
        XSSFWorkbook workbook = new XSSFWorkbook("H:\\→桌面←\\student_info.xlsx");
        // 获取模板sheet
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 找到数据起始行（前两行是标题和表头，要跳过，所以是getRow(2)）
        XSSFRow dataTemplateRow = sheet.getRow(2);
        // 构造一个CellStyle数组，用来存放单元格样式。一行有N个单元格，长度就设置为N
        CellStyle[] cellStyles = new CellStyle[dataTemplateRow.getLastCellNum()];
        for (int i = 0; i < cellStyles.length; i++) {
            // 收集每一个格子对应的格式，你可以理解为准备了一把“格式刷”
            cellStyles[i] = dataTemplateRow.getCell(i).getCellStyle();
        }

        // 创建单元格，并设置样式和数据
        for (int i = 0; i < studentList.size(); i++) {
            // 注意是i+2，模板前两行是大标题和表头。你可能看着难受，想把上面for的i改为i+2，千万别。因为studentList必须从0开始取值
            XSSFRow row = sheet.createRow(i + 2);
            // 为每一行创建单元格并设置数据
            Student student = studentList.get(i);


            XSSFCell nameCell = row.createCell(0);// 创建单元格
            nameCell.setCellStyle(cellStyles[0]);             // 设置单元格样式
            nameCell.setCellValue(student.getName());         // 设置值

            XSSFCell ageCell = row.createCell(1);
            ageCell.setCellStyle(cellStyles[1]);
            ageCell.setCellValue(student.getAge());

            XSSFCell addressCell = row.createCell(2);
            addressCell.setCellStyle(cellStyles[2]);
            addressCell.setCellValue(student.getAddress());

            /**
             * 你可能有疑问，这里是日期类型，是不是要和上一次一样，设置单元格样式为日期类型？
             * 这回不用了，因为上面已经拷贝了模板的样式，生日一栏就是按日期类型展示的
             */
            XSSFCell birthdayCell = row.createCell(3);
            birthdayCell.setCellStyle(cellStyles[3]);
            birthdayCell.setCellValue(student.getBirthday());

            XSSFCell heightCell = row.createCell(4);
            heightCell.setCellStyle(cellStyles[4]);
            heightCell.setCellValue(student.getHeight());

            XSSFCell mainLandChinaCell = row.createCell(5);
            mainLandChinaCell.setCellStyle(cellStyles[5]);
            mainLandChinaCell.setCellValue(student.getIsMainlandChina());
        }

        // 输出
        FileOutputStream out = new FileOutputStream("H:\\→桌面←\\student_info_export_3.xlsx");
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
