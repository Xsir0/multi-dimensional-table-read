package club.xsir.poi;

import org.apache.poi.ss.usermodel.*;

/**
 * @author rwx
 * @version 1.0
 * @description:
 * @date 2025/5/20 10:48
 */
public class Utils {

    public static int calculateColumnWidth(String content, int chineseCharWidth) {
        int width = 0;
        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);
            if (isChinese(c)) {
                width += chineseCharWidth;
            } else {
                width += 2;
            }
        }
        return width;
    }

    public static boolean isChinese(char c) {
        Character.UnicodeBlock ub = Character.UnicodeBlock.of(c);
        return ub == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || ub == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS
                || ub == Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION;
    }

    public static CellStyle renderStyle(Workbook workbook,String fontName,int fontSize,HorizontalAlignment horizontalAlignment,VerticalAlignment verticalAlignment){
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints((short) fontSize);
        font.setBold(false);
        font.setColor(IndexedColors.BLACK.getIndex());
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(horizontalAlignment);
        style.setVerticalAlignment(verticalAlignment);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        return style;
    }

    public static String getLocalPath(){
        return System.getProperty("user.dir");
    }


}
