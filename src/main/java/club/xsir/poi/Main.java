package club.xsir.poi;


import com.alibaba.fastjson.JSON;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * @author rwx
 * @version 1.0
 * @description: TODO
 * @date 2025/3/25 13:54
 */
public class Main {
    public static void main(String[] args) {

        // 对于 XSSF (Excel 2007+)
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

// 合并第1行的A到D列 (行0的列0到3)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

// 写入合并单元格的值
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("合并的标题");

        // 合并A列的第2-5行 (行1到4的列0)
        sheet.addMergedRegion(new CellRangeAddress(1, 4, 0, 0));

// 写入值(只需要在第一行写入)
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("垂直合并");

    }
}