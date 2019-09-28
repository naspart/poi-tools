package com.naspat.poi.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class AbstractExcelExportStyleBuilder {
    private CellStyle tableTitleStyle;
    private CellStyle tableHeaderStyle;
    private CellStyle tableIndexStyle;

    public abstract CellStyle buildTableTitleStyle(Workbook workbook);

    public abstract CellStyle buildTableHeaderStyle(Workbook workbook);

    public abstract CellStyle buildTableIndexStyle(Workbook workbook);

    public final void setWorkbook(Workbook workbook) {
        this.tableTitleStyle = buildTableTitleStyle(workbook);
        this.tableHeaderStyle = buildTableHeaderStyle(workbook);
        this.tableIndexStyle = buildTableIndexStyle(workbook);
    }

    public final CellStyle getTitleStyle() {
        return this.tableTitleStyle;
    }

    public final CellStyle getHeaderStyle() {
        return this.tableHeaderStyle;
    }

    public final CellStyle getIndexStyle() {
        return this.tableIndexStyle;
    }
}
