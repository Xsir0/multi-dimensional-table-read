package club.xsir.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * @author rwx
 * @version 1.0
 * @description:
 * @date 2025/5/20 10:11
 */
public class ExcelWriter {

    public static <T> String writer(String filePath,Class<T> dataType, BiConsumer<Workbook,Sheet> consumer){
        Map<Integer, Field> fieldsInOrderMap = FieldOrderUtil.getFieldsInOrderMap(dataType);
        Workbook workbook = new XSSFWorkbook();

        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        font.setColor(IndexedColors.ROSE.getIndex());
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        Sheet sheet = workbook.createSheet("Sheet1");
        sheet.createFreezePane(0, 1);
        Row titleRow = sheet.createRow(0);
        for (Integer i : fieldsInOrderMap.keySet()) {
            FieldInfo order = fieldsInOrderMap.get(i).getAnnotation(FieldInfo.class);
            Cell titleCell = titleRow.createCell(i-1);
            titleCell.setCellStyle(style);
            sheet.setColumnWidth(i-1,Utils.calculateColumnWidth(order.name(),3)*256);
            titleCell.setCellValue(order.name());
        }
        consumer.accept(workbook,sheet);
        try (FileOutputStream out = new FileOutputStream(filePath)) {
            workbook.write(out);
            workbook.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return null;
    }

}
