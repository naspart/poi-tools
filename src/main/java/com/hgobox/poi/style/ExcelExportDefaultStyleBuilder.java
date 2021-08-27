package com.hgobox.poi.style;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

public class ExcelExportDefaultStyleBuilder extends AbstractExcelExportStyleBuilder {
    @Override
    public CellStyle buildTableHeaderStyle(Workbook workbook) {
        Font headFont = workbook.createFont();
        headFont.setFontName("宋体");
        headFont.setFontHeightInPoints((short) 11);
        headFont.setCharSet(Font.DEFAULT_CHARSET);
        headFont.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFont(headFont);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return cellStyle;
    }

    @Override
    public CellStyle buildTableTitleStyle(Workbook workbook) {
        Font titleFont = workbook.createFont();
        titleFont.setFontName("华文楷体");
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setCharSet(Font.DEFAULT_CHARSET);
        titleFont.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        cellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        cellStyle.setBorderRight(BorderStyle.THIN);// 右边框
        cellStyle.setFont(titleFont);       // 设置字体
        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.DARK_TEAL.getIndex());   // 设置背景颜色
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);     // 设置背景填充方式

        return cellStyle;
    }

    @Override
    public CellStyle buildTableIndexStyle(Workbook workbook) {
        Font rowNumFont = workbook.createFont();
        rowNumFont.setFontName("华文楷体");
        rowNumFont.setFontHeightInPoints((short) 10);
        rowNumFont.setCharSet(Font.DEFAULT_CHARSET);
        rowNumFont.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());

        CellStyle rowNumStyle = workbook.createCellStyle();
        rowNumStyle.setAlignment(HorizontalAlignment.CENTER);
        rowNumStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        rowNumStyle.setBorderTop(BorderStyle.THIN);
        rowNumStyle.setBorderBottom(BorderStyle.THIN);
        rowNumStyle.setBorderLeft(BorderStyle.THIN);
        rowNumStyle.setBorderRight(BorderStyle.THIN);
        rowNumStyle.setFont(rowNumFont);

        rowNumStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getIndex());
        rowNumStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return rowNumStyle;
    }
}
