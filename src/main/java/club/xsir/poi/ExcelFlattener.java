package club.xsir.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelFlattener {
    public static void main(String[] args) throws IOException, InvalidFormatException, InstantiationException, IllegalAccessException {
        String filePath = "demo.xls";
        List<DataStruct> generate = generate(filePath,DataStruct.class);
        for (DataStruct dataStructs : generate) {
            System.out.println(dataStructs);
            System.out.println("=========================================================");
        }
        getLocalPath();
    }

    public static <T>  List<T>  generate(String filePath,Class<T> dataType) throws IOException, InvalidFormatException, InstantiationException, IllegalAccessException {

        Map<Integer, Field> fieldsInOrderMap = FieldOrderUtil.getFieldsInOrderMap(dataType);

        if (fieldsInOrderMap==null || fieldsInOrderMap.isEmpty()){
            throw new RuntimeException("实体类设置错误!");
        }

        List<List<Object>> flattenedData = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(new File(filePath))) {
            Sheet sheet = workbook.getSheetAt(0);
            Row _row = sheet.getRow(0);
            short numColumns = _row.getLastCellNum();
            if ((numColumns-_row.getFirstCellNum())!=fieldsInOrderMap.size()){
                throw new RuntimeException("实体类配置字段数跟 Excel 不一致");
            }

            // 获取所有合并区域
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (Row row : sheet) {
                List<Object> rowData = new ArrayList<>();
                for (int i = 0; i < numColumns; i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    CellRangeAddress mergedRegion = getMergedRegion(mergedRegions,
                            cell.getRowIndex(), cell.getColumnIndex());

                    if (mergedRegion != null) {
                        // 如果是合并区域且不是第一个单元格，跳过
                        if (cell.getRowIndex() != mergedRegion.getFirstRow() ||
                                cell.getColumnIndex() != mergedRegion.getFirstColumn()) {
                            continue;
                        }
                    }
                    rowData.add(getCellValue(cell));
                }
                //for (Cell cell : row) {
                //    CellRangeAddress mergedRegion = getMergedRegion(mergedRegions,
                //            cell.getRowIndex(), cell.getColumnIndex());
                //
                //    if (mergedRegion != null) {
                //        // 如果是合并区域且不是第一个单元格，跳过
                //        if (cell.getRowIndex() != mergedRegion.getFirstRow() ||
                //                cell.getColumnIndex() != mergedRegion.getFirstColumn()) {
                //            continue;
                //        }
                //    }
                //    rowData.add(getCellValue(cell));
                //}
                flattenedData.add(rowData);
            }

            for (CellRangeAddress region : mergedRegions) {
                if (flattenedData.size()<=region.getFirstRow() || flattenedData.get(region.getFirstRow()).size()<=region.getFirstColumn()){
                    continue;
                }
                Object value = flattenedData.get(region.getFirstRow()).get(region.getFirstColumn());

                // 填充合并的行
                for (int r = region.getFirstRow() + 1; r <= region.getLastRow(); r++) {
                    if (r < flattenedData.size()) {
                        flattenedData.get(r).add(region.getFirstColumn(), value);
                    }
                }

                // 填充合并的列
                for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) {
                    if (r < flattenedData.size()) {
                        for (int c = region.getFirstColumn() + 1; c <= region.getLastColumn(); c++) {
                            if (c < flattenedData.get(r).size()) {
                                flattenedData.get(r).set(c, value);
                            }
                        }
                    }
                }
            }
        }
        List<T> result = new ArrayList<>();
        for (List<Object> flattenedDatum : flattenedData) {
            T t = dataType.newInstance();
            for (int i = 0; i < flattenedDatum.size(); i++) {
                Object fieldValue = flattenedDatum.get(i);
                if (null != fieldValue && !fieldValue.toString().trim().isEmpty()){
                    Field field = fieldsInOrderMap.get(i+1);
                    field.setAccessible(true);
                    field.set(t,fieldValue.toString());
                    field.setAccessible(false);
                }
            }
            result.add(t);
        }
        return result;
    }

    private static CellRangeAddress getMergedRegion(List<CellRangeAddress> regions,
            int row, int col) {
        for (CellRangeAddress region : regions) {
            if (region.isInRange(row, col)) {
                return region;
            }
        }
        return null;
    }

    private static Object getCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN: return cell.getBooleanCellValue();
            case FORMULA:
                switch (cell.getCachedFormulaResultTypeEnum()) {
                    case STRING: return cell.getStringCellValue();
                    case NUMERIC: return cell.getNumericCellValue();
                    case BOOLEAN: return cell.getBooleanCellValue();
                    default: return cell.getCellFormula();
                }
            default: return "";
        }
    }

    private static void getLocalPath(){
        String userDir = System.getProperty("user.dir");
        System.out.println(userDir);
    }
}