package com.naspart.poi.style;

import org.apache.poi.ss.usermodel.CellStyle;

public interface IExcelExportStyleBuilder {
    /**
     * 列表头样式
     */
    CellStyle getHeaderStyle();

    /**
     * 标题样式
     */
    CellStyle getTitleStyle();

    /**
     * 序号列样式
     */
    CellStyle getIndexStyle();
}