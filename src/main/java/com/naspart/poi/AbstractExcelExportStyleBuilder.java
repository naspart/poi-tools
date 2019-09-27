package com.naspart.poi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class AbstractExcelExportStyleBuilder implements IExcelExportStyleBuilder {
    protected Workbook workbook;
    private CellStyle tableTitleStyle;
    private CellStyle tableHeaderStyle;
    private CellStyle tableIndexStyle;

    protected AbstractExcelExportStyleBuilder(Workbook workbook) {
        this.workbook = workbook;
        this.tableTitleStyle = getTableTitleStyle();
        this.tableHeaderStyle = getTableHeaderStyle();
        this.tableIndexStyle = getTableIndexStyle();
    }

    public abstract CellStyle getTableTitleStyle();

    public abstract CellStyle getTableHeaderStyle();

    public abstract CellStyle getTableIndexStyle();

    @Override
    public CellStyle getTitleStyle() {
        return this.tableTitleStyle;
    }

    @Override
    public CellStyle getHeaderStyle() {
        return this.tableHeaderStyle;
    }

    @Override
    public CellStyle getIndexStyle() {
        return this.tableIndexStyle;
    }
}
