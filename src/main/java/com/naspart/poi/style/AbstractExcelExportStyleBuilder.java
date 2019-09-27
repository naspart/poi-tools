package com.naspart.poi.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class AbstractExcelExportStyleBuilder implements IExcelExportStyleBuilder {
    protected Workbook workbook;
    private CellStyle tableTitleStyle;
    private CellStyle tableHeaderStyle;
    private CellStyle tableIndexStyle;

    protected AbstractExcelExportStyleBuilder(Workbook workbook) {
        this.workbook = workbook;
        this.tableTitleStyle = buildTableTitleStyle();
        this.tableHeaderStyle = buildTableHeaderStyle();
        this.tableIndexStyle = buildTableIndexStyle();
    }

    public abstract CellStyle buildTableTitleStyle();

    public abstract CellStyle buildTableHeaderStyle();

    public abstract CellStyle buildTableIndexStyle();

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
