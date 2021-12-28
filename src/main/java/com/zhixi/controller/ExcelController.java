package com.zhixi.controller;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author zhangzhixi
 * @version 1.0
 * @date 2021-12-28 16:36
 */
@RestController
public class ExcelController {

    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    @GetMapping("/exportExcel")
    public void exportExcel(HttpServletResponse response) throws Exception {
        // 模拟从数据库查询数据
        List<Student> studentList = new ArrayList<>();
        studentList.add(new Student(1L, "周深（web导出）", 28, "贵州", new SimpleDateFormat("yyyy-MM-dd").parse("1992-9-29"), 161.0, true));
        studentList.add(new Student(2L, "李健（web导出）", 46, "哈尔滨", new SimpleDateFormat("yyyy-MM-dd").parse("1974-9-23"), 174.5, true));
        studentList.add(new Student(3L, "周星驰（web导出）", 58, "香港", new SimpleDateFormat("yyyy-MM-dd").parse("1962-6-22"), 174.0, false));

        // 读取模板（实际开发可以放在resources文件夹下，随着项目一起打包发布）
        InputStream excelInputStream = new ClassPathResource("static/excel/student_info.xlsx").getInputStream();
        // XSSFWorkbook除了直接接收Path外，还可以传入输入流
        XSSFWorkbook workbook = new XSSFWorkbook(excelInputStream);
        // 获取模板sheet
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 找到数据起始行（前两行是标题和表头，要跳过，所以是getRow(2)）
        XSSFRow dataTemplateRow = sheet.getRow(2);
        // 构造一个CellStyle数组，用来存放单元格样式。一行有N个单元格，数组初始长度就设置为N
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
            nameCell.setCellValue(student.getName());         // 设置值
            nameCell.setCellStyle(cellStyles[0]);             // 设置单元格样式

            XSSFCell ageCell = row.createCell(1);
            ageCell.setCellValue(student.getAge());
            ageCell.setCellStyle(cellStyles[1]);

            XSSFCell addressCell = row.createCell(2);
            addressCell.setCellValue(student.getAddress());
            addressCell.setCellStyle(cellStyles[2]);

            /**
             * 你可能有疑问，这里是日期类型，是不是要和上一次一样，设置单元格样式为日期类型？
             * 这回不用了，因为上面已经拷贝了模板的样式，生日一栏就是按日期类型展示的
             */
            XSSFCell birthdayCell = row.createCell(3);
            birthdayCell.setCellValue(student.getBirthday());
            birthdayCell.setCellStyle(cellStyles[3]);

            XSSFCell heightCell = row.createCell(4);
            heightCell.setCellValue(student.getHeight());
            heightCell.setCellStyle(cellStyles[4]);

            XSSFCell mainLandChinaCell = row.createCell(5);
            mainLandChinaCell.setCellValue(student.getIsMainlandChina());
            mainLandChinaCell.setCellStyle(cellStyles[5]);
        }

        /**
         * 之前通过本地文件流输出到桌面：
         * FileOutputStream out = new FileOutputStream("/Users/kevin/Documents/study/student_info_export.xlsx");
         * 现在用网络流：response.getOutputStream()
         * 注意，response的响应流没必要手动关闭，交给Tomcat关闭
         */
        String fileName = new String("学生信息表.xlsx".getBytes("UTF-8"), "ISO-8859-1");
        response.setContentType("application/octet-stream");
        response.setHeader("content-disposition", "attachment;filename=" + fileName);
        response.setHeader("filename", fileName);
        workbook.write(response.getOutputStream());
        workbook.close();
        logger.info("导出学生信息表成功！");
    }

    @PostMapping("/importExcel")
    public Map importExcel(MultipartFile file) throws Exception {
        // 直接获取上传的文件流，传入构造函数
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        // 获取工作表。一个工作薄中可能有多个工作表，比如sheet1 sheet2，可以根据下标，也可以根据sheet名称。这里根据下标即可。
        XSSFSheet sheet = workbook.getSheetAt(0);

        // 收集每一行数据（跳过标题和表头，所以int i = 2）
        int lastRowNum = sheet.getLastRowNum();
        List<Student> studentList = new ArrayList<>();
        for (int i = 2; i <= lastRowNum; i++) {
            // 收集当前行所有单元格的数据
            XSSFRow row = sheet.getRow(i);
            short lastCellNum = row.getLastCellNum();
            List<String> cellDataList = new ArrayList<>();
            for (int j = 0; j < lastCellNum; j++) {
                cellDataList.add(getValue(row.getCell(j)));
            }

            // 把当前行数据设置到POJO。由于Excel单元格的顺序和POJO字段顺序一致，也就是数据类型一致，所以可以直接强转
            Student student = new Student();
            student.setName(cellDataList.get(0));
            student.setAge(Integer.parseInt(cellDataList.get(1)));
            student.setAddress(cellDataList.get(2));
            // getValue()方法返回的是字符串类型的 1962-6-22 00:00:00，这里按"yyyy-MM-dd HH:mm:ss"重新解析为Date
            student.setBirthday(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(cellDataList.get(3)));
            student.setHeight(Double.parseDouble(cellDataList.get(4)));
            student.setHeight(Double.parseDouble(cellDataList.get(4)));
            student.setIsMainlandChina(Boolean.valueOf(cellDataList.get(5)));
            studentList.add(student);
        }

        // 插入数据库
        saveToDB(studentList);
        logger.info("导入{}成功！", file.getOriginalFilename());

        Map<String, Object> result = new HashMap<>();
        result.put("code", 200);
        result.put("data", null);
        result.put("msg", "success");
        return result;
    }

    private void saveToDB(List<Student> studentList) {
        if (CollectionUtils.isEmpty(studentList)) {
            return;
        }
        // 直接打印，模拟插入数据库
        studentList.forEach(System.out::println);
    }

    /**
     * 提供POI数据类型 --> Java数据类型的转换
     * 由于本方法返回值设为String，所以不管转换后是什么Java类型，都要以String格式返回
     * 所以Date会被格式化为yyyy-MM-dd HH:mm:ss
     * 后面根据需要自己另外转换
     *
     * @param cell
     * @return
     */
    private String getValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().getString().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // DateUtil是POI内部提供的日期工具类，可以把原本是日期类型的NUMERIC转为Java的Data类型
                    Date javaDate = DateUtil.getJavaDate(cell.getNumericCellValue());
                    String dateString = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(javaDate);
                    return dateString;
                } else {
                    /*
                     * 无论Excel中是58还是58.0，数值类型在POI中最终都被解读为Double。
                     * 这里的解决办法是通过BigDecimal先把Double先转成字符串，如果是.0结尾，把.0去掉
                     * */
                    String strCell = "";
                    Double num = cell.getNumericCellValue();
                    BigDecimal bd = new BigDecimal(num.toString());
                    if (bd != null) {
                        strCell = bd.toPlainString();
                    }
                    // 去除 浮点型 自动加的 .0
                    if (strCell.endsWith(".0")) {
                        strCell = strCell.substring(0, strCell.indexOf("."));
                    }
                    return strCell;
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    @Data
    @NoArgsConstructor
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