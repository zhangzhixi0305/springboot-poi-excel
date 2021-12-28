package com.zhixi.test;

/**
 * @author zhangzhixi
 * @version 1.0
 * @date 2021-12-28 14:23
 */

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 最简单的Excel测试
 * 目前为止，我们只是引入了POI依赖，然后准备了一个Excel文件
 * 下面写了两段代码，大家花时间阅读一下
 */
public class SimpleTest {
    /*@Test
    public void testSimpleRead() throws IOException {
        // 获取工作薄（把路径换成你本地的）
        XSSFWorkbook workbook = new XSSFWorkbook("H:\\→桌面←\\student_info.xlsx");
        // 获取工作表。一个工作薄中可能有多个工作表，比如sheet1 sheet2，可以根据下标，也可以根据sheet名称
        XSSFSheet sheet = workbook.getSheetAt(0);

        *//**
     * 方式1 增强for，Row代表行，Cell代表单元格
     *//*
        // 遍历每一行
        for (Row row : sheet) {
            // 遍历当前行的每个单元格
            for (Cell cell : row) {
                // 获取单元格类型
                CellType cellType = cell.getCellType();
                // 根据类型选择匹配的getXxxValue()方法，比如你判断当前单元格的值是BOOLEAN类型的，你就要用getBooleanCellValue()
                if (cellType == CellType.NUMERIC) {
                    // 你可以试着把这里的getNumericCellValue()改成getStringCellValue()，观察会发生什么
                    System.out.print(cell.getNumericCellValue() + "\t");
                } else if (cellType == CellType.BOOLEAN) {
                    System.out.print(cell.getBooleanCellValue() + "\t");
                } else if (cellType == CellType.STRING) {
                    System.out.print(cell.getStringCellValue() + "\t");
                }
            }
            System.out.println("\n------------------------------------------");
        }

        System.out.println("\n==========================================\n");

        */

    /**
     * 方式2 普通for
     *//*
        // 获取最后一行
        int lastRowNum = sheet.getLastRowNum();
        // 从第一行开始遍历，直到最后一行
        for (int j = 0; j <= lastRowNum; j++) {
            // 获取当前行
            XSSFRow row = sheet.getRow(j);
            if (row != null) {
                // 获取当前行最后一个单元格
                short cellNum = row.getLastCellNum();
                // 从第一个单元格开始遍历，直到最后一个单元格
                for (int k = 0; k < cellNum; k++) {
                    XSSFCell cell = row.getCell(k);
                    CellType cellType = cell.getCellType();
                    // 注意，我们的EXCEL有NUMERIC、STRING和BOOLEAN三种类型，但这里省略了BOOLEAN，会发生什么呢？
                    if (cellType == CellType.NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + "\t");
                    } else if (cellType == CellType.STRING) {
                        System.out.print(cell.getStringCellValue() + "\t");
                    }
                }
            }
            System.out.println("\n------------------------------------------");
        }
    }*/
    @Test
    public void testSimpleRead() throws IOException {
        // 获取工作薄（把路径换成你本地的）
        XSSFWorkbook workbook = new XSSFWorkbook("H:\\→桌面←\\student_info.xlsx");
        // 获取工作表。一个工作薄中可能有多个工作表，比如sheet1 sheet2，可以根据下标，也可以根据sheet名称
        XSSFSheet sheet = workbook.getSheetAt(0);

        // 这里采用普通for
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                short cellNum = row.getLastCellNum();
                for (int k = 0; k < cellNum; k++) {
                    XSSFCell cell = row.getCell(k);
                    // 要判断类型并采用对应的get方法，由于比较繁琐，我们抽取成方法
                    System.out.print(getValue(cell) + "\t");
                }
            }
            System.out.println("\n------------------------------------------");
        }

    }

    /**
     * 提供POI数据类型 到 Java数据类型的转换，最终都返回String
     *
     * @param cell
     * @return
     */
    private String getValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        // 常用的一般就这三大类：STRING、NUMERIC、BOOLEAN，几乎没有别的类型了。但NUMERIC要细分，特别注意
        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().getString().trim();

            // EXCEL的日期和数字都被POI整合为NUMERIC，这里把它们重新拆开
            case NUMERIC:
                // DateUtil是POI内部提供的日期工具类，可以把原本是日期类型的NUMERIC转为Java的Data类型
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date javaDate = DateUtil.getJavaDate(cell.getNumericCellValue());
                    return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(javaDate);
                } else {
                    // 无论EXCEL中是58还是58.0，数值类型在POI中最终都被解读为double。这里的解决办法是通过BigDecimal先把double先转成字符串，如果是.0结尾，把.0去掉
                    String strCell = "";
                    double num = cell.getNumericCellValue();
                    BigDecimal bd = new BigDecimal(Double.toString(num));
                    strCell = bd.toPlainString();
                    // 去除 浮点型 自动加的 .0
                    if (strCell.endsWith(".0")) {
                        strCell = strCell.substring(0, strCell.indexOf("."));
                    }
                    return strCell;
                }
            case BOOLEAN:
                boolean booleanCellValue = cell.getBooleanCellValue();
                return String.valueOf(booleanCellValue);
            default:
                return "";
        }
    }
}
