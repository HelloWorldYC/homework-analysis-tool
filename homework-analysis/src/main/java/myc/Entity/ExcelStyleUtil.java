package myc.Entity;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;


public class ExcelStyleUtil {

    private Font firstHeaderFont;
    private CellStyle firstHeaderStyle;
    private Font secondHeaderFont;
    private CellStyle secondHeaderStyle;
    private Font dataFont;
    private CellStyle dataCellStyle;
    private Font summarizationRowFont;
    private CellStyle summarizationRowStyle;

    public Font getFirstHeaderFont() {
        return firstHeaderFont;
    }
    public CellStyle getFirstHeaderStyle() {
        return firstHeaderStyle;
    }
    public Font getSecondHeaderFont() {
        return secondHeaderFont;
    }
    public CellStyle getSecondHeaderStyle() {
        return  secondHeaderStyle;
    }
    public Font getDataFont() {
        return dataFont;
    }
    public CellStyle getDataCellStyle(){
        return dataCellStyle;
    }
    public Font getSummarizationRowFont(){
        return summarizationRowFont;
    }
    public CellStyle getSummarizationRowStyle(){
        return summarizationRowStyle;
    }

    public ExcelStyleUtil(Workbook workbook) {
        // 首行表头样式
        firstHeaderFont = workbook.createFont();
        firstHeaderFont.setFontName("微软雅黑");
        firstHeaderFont.setBold(true); // 加粗
        firstHeaderFont.setFontHeightInPoints((short) 20);
        firstHeaderStyle = workbook.createCellStyle();
        firstHeaderStyle.setFont(firstHeaderFont);
        firstHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        firstHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        firstHeaderStyle.setWrapText(true);

        // 第二行表头样式
        secondHeaderFont = workbook.createFont();
        secondHeaderFont.setFontName("微软雅黑");
        secondHeaderFont.setFontHeightInPoints((short) 16);
        secondHeaderStyle = workbook.createCellStyle();
        secondHeaderStyle.setFont(secondHeaderFont);
        secondHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        secondHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        secondHeaderStyle.setWrapText(true);

        // 表数据样式
        dataFont = workbook.createFont();
        dataFont.setFontName("宋体");
        dataFont.setFontHeightInPoints((short) 16);
        dataCellStyle = workbook.createCellStyle();
        dataCellStyle.setFont(dataFont);
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        DataFormat dataFormat = workbook.createDataFormat();
        dataCellStyle.setDataFormat(dataFormat.getFormat("#0.00")); // 设置格式，保留两位小数
        dataCellStyle.setWrapText(true);

        // 汇总行样式
        summarizationRowFont = workbook.createFont();
        summarizationRowFont.setFontName("微软雅黑");
        summarizationRowFont.setFontHeightInPoints((short) 16);
        summarizationRowStyle = workbook.createCellStyle();
        summarizationRowStyle.setFont(summarizationRowFont);
        summarizationRowStyle.setAlignment(HorizontalAlignment.CENTER);
        summarizationRowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        summarizationRowStyle.setDataFormat(dataFormat.getFormat("#0.00")); // 设置格式，保留两位小数
        summarizationRowStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex()); // 设置单元格背景色为绿色
        summarizationRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        summarizationRowStyle.setWrapText(true);
    }


}
