package com.naspat.poi;

import com.naspat.poi.annotation.ExcelField;
import com.naspat.poi.annotation.ExcelTarget;
import com.naspat.poi.entity.ExcelFieldEntity;
import com.naspat.poi.entity.ExcelTargetEntity;
import com.naspat.poi.enums.VerticalAlignment;
import com.naspat.poi.function.ExcelDataFunction;
import com.naspat.poi.style.AbstractExcelExportStyleBuilder;
import com.naspat.poi.style.ExcelExportDefaultStyleBuilder;
import com.naspat.poi.util.BeanUtils;
import com.naspat.poi.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Slf4j
public class ExcelExportBuilder {
    private final SXSSFWorkbook workbook;
    private AbstractExcelExportStyleBuilder excelExportStyleBuilder;

    public ExcelExportBuilder() {
        this.workbook = new SXSSFWorkbook(200);
        this.excelExportStyleBuilder = new ExcelExportDefaultStyleBuilder();
        this.excelExportStyleBuilder.setWorkbook(workbook);

        log.debug("使用默认样式构造器...");
    }

    public ExcelExportBuilder(AbstractExcelExportStyleBuilder excelExportStyle) {
        this.workbook = new SXSSFWorkbook();
        this.excelExportStyleBuilder = excelExportStyle;
        this.excelExportStyleBuilder.setWorkbook(workbook);

        log.debug("使用自定义样式构造器...");
    }

    public void setExcelExportStyleBuilder(AbstractExcelExportStyleBuilder excelExportStyleBuilder) {
        this.excelExportStyleBuilder = excelExportStyleBuilder;
        this.excelExportStyleBuilder.setWorkbook(workbook);
    }

    public <R, T> ExcelExportBuilder build(String sheetName, R param, ExcelDataFunction<R, T> excelDataFunction, Class<?> clazz) {
        return this.build(null, sheetName, param, excelDataFunction, clazz);
    }

    public <R, T> ExcelExportBuilder build(String sheetTitle, String sheetName, R param, ExcelDataFunction<R, T> excelDataFunction, Class<?> clazz) {
        Sheet sheet = workbook.createSheet(sheetName);

        ExcelTargetEntity excelTargetEntity = this.getExcelTargetEntity(sheet, clazz);

        if (StringUtils.isNotBlank(sheetTitle)) {
            excelTargetEntity.setSheetTitle(sheetTitle);
        }

        this.initTableTitle(sheet, excelTargetEntity);
        this.initTableHeader(sheet, excelTargetEntity);

        int firstPageNo = 1;
        int rowNum = excelTargetEntity.isHasSheetTitle() ? 2 : 1;
        while (true) {
            Collection<T> dataSet = excelDataFunction.getData(param, firstPageNo, 3000);
            if (dataSet == null || dataSet.isEmpty()) {
                break;
            }

            for (T data : dataSet) {
                Object result = excelDataFunction.convert(data);
                this.createRow(sheet, excelTargetEntity, result, rowNum);

                rowNum++;
            }

            firstPageNo++;
        }

        sheet.createFreezePane(excelTargetEntity.getFrozenColumns(), excelTargetEntity.getFrozenRows());

        return this;
    }

    public ExcelExportBuilder build(String sheetName, Collection<?> dataSet, Class<?> clazz) {
        return this.build(null, sheetName, dataSet, clazz);
    }

    public ExcelExportBuilder build(String sheetTitle, String sheetName, Collection<?> dataSet, Class<?> clazz) {
        Sheet sheet = workbook.createSheet(sheetName);

        ExcelTargetEntity excelTargetEntity = this.getExcelTargetEntity(sheet, clazz);

        if (StringUtils.isNotBlank(sheetTitle)) {
            excelTargetEntity.setSheetTitle(sheetTitle);
        }

        this.initTableTitle(sheet, excelTargetEntity);
        this.initTableHeader(sheet, excelTargetEntity);

        int rowNum = excelTargetEntity.isHasSheetTitle() ? 2 : 1;
        for (Object data : dataSet) {
            this.createRow(sheet, excelTargetEntity, data, rowNum);

            rowNum++;
        }

        sheet.createFreezePane(excelTargetEntity.getFrozenColumns(), excelTargetEntity.getFrozenRows());

        return this;
    }

    public void write(OutputStream outputStream) throws IOException {
        try {
            if (outputStream != null) {
                this.workbook.write(outputStream);
                outputStream.flush();
            }
        } finally {
            this.workbook.dispose();
            this.workbook.close();
        }
    }

    private ExcelTargetEntity getExcelTargetEntity(Sheet sheet, Class<?> clazz) {
        Field[] fields = BeanUtils.getClassFields(clazz);

        ExcelTarget excelTarget = clazz.getAnnotation(ExcelTarget.class);
        String sheetTitle = (excelTarget == null ? "" : excelTarget.title());
        int frozenColumns = excelTarget == null ? 0 : excelTarget.frozenColumns();
        int frozenRows = excelTarget == null ? 0 : excelTarget.frozenRows();

        boolean hasSheetTitle = StringUtils.isNotBlank(sheetTitle);

        List<Field> fieldList = new ArrayList<>();
        Map<Field, ExcelFieldEntity> excelFieldEntityMap = new HashMap<>();
        Map<Field, CellStyle> cellStyleMap = new HashMap<>();
        for (int index = 0; index < fields.length; index++) {
            Field field = fields[index];
            fieldList.add(field);

            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null) {
                CellStyle cellStyle = getCellStyle(sheet, excelField.format(), excelField.horizontalAlignment(), excelField.verticalAlignment());
                sheet.setColumnWidth(index, excelField.width() * 256);

                Map<String, String> replaceMap = new HashMap<>();
                String[] replace = excelField.replace();
                for (String s : replace) {
                    String[] tmp = s.split("_");
                    replaceMap.put(tmp[0], tmp[1]);
                }

                excelFieldEntityMap.put(
                        field, ExcelFieldEntity.builder()
                                .field(field)
                                .name(excelField.name())
                                .replaceMap(replaceMap)
                                .format(excelField.format())
                                .width(excelField.width())
                                .horizontalAlignment(excelField.horizontalAlignment())
                                .verticalAlignment(excelField.verticalAlignment())
                                .verticalMerge(excelField.verticalMerge())
                                .build()
                );

                cellStyleMap.put(field, cellStyle);
            }
        }

        return ExcelTargetEntity.builder()
                .hasSheetTitle(hasSheetTitle)
                .sheetTitle(sheetTitle)
                .frozenColumns(frozenColumns)
                .frozenRows(frozenRows)
                .fields(fieldList)
                .excelFieldEntityMap(excelFieldEntityMap)
                .cellStyleMap(cellStyleMap)
                .build();
    }

    private CellStyle getCellStyle(Sheet sheet, String format, com.naspat.poi.enums.HorizontalAlignment horizontalAlignment, com.naspat.poi.enums.VerticalAlignment verticalAlignment) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        // 设置水平对齐方式
        if (horizontalAlignment == com.naspat.poi.enums.HorizontalAlignment.CENTER) {
            cellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        } else if (horizontalAlignment == com.naspat.poi.enums.HorizontalAlignment.RIGHT) {
            cellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
        } else if (horizontalAlignment == com.naspat.poi.enums.HorizontalAlignment.LEFT) {
            cellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        } else {
            cellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL);
        }

        // 设置垂直对齐方式
        if (verticalAlignment == VerticalAlignment.TOP) {
            cellStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.TOP);
        } else if (verticalAlignment == VerticalAlignment.CENTER) {
            cellStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        } else if (verticalAlignment == VerticalAlignment.BOTTOM) {
            cellStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM);
        }

        cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        cellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        cellStyle.setBorderRight(BorderStyle.THIN);// 右边框

        // 使用自定义的日期格式
        DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(format));

        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Courier New");
        font.setFontHeightInPoints((short) 10);
        font.setCharSet(Font.DEFAULT_CHARSET);
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());

        // 设置字体
        cellStyle.setFont(font);
        // 设置背景颜色
        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        // 设置背景填充方式
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return cellStyle;
    }

    private void initTableTitle(Sheet sheet, ExcelTargetEntity entity) {
        if (StringUtils.isNotBlank(entity.getSheetTitle())) {
            CellStyle tableTitleStyle = excelExportStyleBuilder.getTitleStyle();

            CellRangeAddress titleRange = new CellRangeAddress(0, 0, 0, entity.getExcelFieldEntityMap().size() - 1);
            sheet.addMergedRegion(titleRange);

            Row titleRow = sheet.createRow(0);
            titleRow.setHeight((short) 800);

            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(tableTitleStyle);
            titleCell.setCellValue(entity.getSheetTitle());
        }
    }

    private void initTableHeader(Sheet sheet, ExcelTargetEntity entity) {
        CellStyle tableHeaderStyle = excelExportStyleBuilder.getHeaderStyle();

        Row headRow = sheet.createRow(entity.isHasSheetTitle() ? 1 : 0);
        headRow.setHeight((short) 350);

        for (int index = 0; index < entity.getFields().size(); index++) {
            Field field = entity.getFields().get(index);
            ExcelFieldEntity tmp = entity.getExcelFieldEntityMap().get(field);

            Cell headCell = headRow.createCell(index);
            headCell.setCellStyle(tableHeaderStyle);
            headCell.setCellValue(tmp.getName());
        }
    }

    private void createRow(Sheet sheet, ExcelTargetEntity excelTargetEntity, Object obj, int curRowIndex) {
        Row currentRow = sheet.createRow(curRowIndex);
        Row lastRow = sheet.getRow(curRowIndex - 1);
        currentRow.setHeight((short) 300);

        Cell[] cells = this.createCells(currentRow, excelTargetEntity.getFields().size());
        int cellNum = 0;
        for (Field field : excelTargetEntity.getFields()) {
            Object value = BeanUtils.getFieldValue(obj, field.getName());
            ExcelFieldEntity entity = excelTargetEntity.getExcelFieldEntityMap().get(field);

            if (value == null) {
                cells[cellNum].setCellValue("");
            } else {
                Class<?> type = field.getType();
                switch (type.getName()) {
                    case "int":
                    case "java.lang.Integer":
                        if (!entity.getReplaceMap().isEmpty()) {
                            String val = entity.getReplaceMap().get(value.toString());
                            if (val == null) {
                                cells[cellNum].setCellValue((int) value);
                            } else {
                                cells[cellNum].setCellValue(val);
                            }
                        } else {
                            cells[cellNum].setCellValue((int) value);
                        }
                        break;

                    case "long":
                    case "java.lang.Long":
                        if (!entity.getReplaceMap().isEmpty()) {
                            String val = entity.getReplaceMap().get(value.toString());
                            if (val == null) {
                                cells[cellNum].setCellValue((long) value);
                            } else {
                                cells[cellNum].setCellValue(val);
                            }
                        } else {
                            cells[cellNum].setCellValue((long) value);
                        }
                        break;

                    case "float":
                    case "java.lang.Float":
                        cells[cellNum].setCellValue((float) value);
                        break;

                    case "double":
                    case "java.lang.Double":
                        cells[cellNum].setCellValue((double) value);
                        break;

                    case "java.util.Date":
                        cells[cellNum].setCellValue((Date) value);
                        break;

                    case "java.util.Calendar":
                        cells[cellNum].setCellValue((Calendar) value);
                        break;

                    case "java.time.LocalTime":
                        cells[cellNum].setCellValue(((LocalTime) value).format(DateTimeFormatter.ofPattern(entity.getFormat())));
                        break;

                    case "java.time.LocalDate":
                        cells[cellNum].setCellValue(Date.from(((LocalDate) value).atStartOfDay(ZoneId.systemDefault()).toInstant()));
                        break;

                    case "java.time.LocalDateTime":
                        cells[cellNum].setCellValue(Date.from(((LocalDateTime) value).atZone(ZoneId.systemDefault()).toInstant()));
                        break;

                    default:
                        cells[cellNum].setCellValue(value.toString());
                        break;

                }
            }

            cells[cellNum].setCellStyle(excelTargetEntity.getCellStyleMap().get(entity.getField()));

            if (entity.isVerticalMerge()) {
                String currentValue = this.getCellValueByCell(currentRow.getCell(excelTargetEntity.getMergeReferenceIndex()));
                String lastValue = this.getCellValueByCell(lastRow.getCell(excelTargetEntity.getMergeReferenceIndex()));
                if (currentValue.equals(lastValue)) {
                    List<CellRangeAddress> mergeRegions = sheet.getMergedRegions();
                    boolean isMerged = false;
                    for (int i = 0; i < mergeRegions.size() && !isMerged; i++) {
                        CellRangeAddress cellRangeAddr = mergeRegions.get(i);
                        // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
                        if (cellRangeAddr.isInRange(curRowIndex - 1, cellNum)) {
                            sheet.removeMergedRegion(i);
                            cellRangeAddr.setLastRow(curRowIndex);
                            sheet.addMergedRegion(cellRangeAddr);
                            isMerged = true;
                        }
                    }

                    // 若上一个单元格未被合并，则新增合并单元
                    if (!isMerged) {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(curRowIndex - 1, curRowIndex, cellNum, cellNum);
                        sheet.addMergedRegion(cellRangeAddress);
                    }
                }
            }

            cellNum++;
        }
    }

    private Cell[] createCells(Row row, int num) {
        Cell[] cells = new Cell[num];
        for (int i = 0, len = cells.length; i < len; i++) {
            cells[i] = row.createCell(i);
        }

        return cells;
    }

    private String getCellValueByCell(Cell cell) {
        if (cell == null || cell.toString().trim().equals("")) {
            return "";
        }

        String cellValue = "";
        CellType cellType = cell.getCellType();
        if (cellType.equals(CellType.STRING)) {
            cellValue = cell.getStringCellValue().trim();
        } else if (cellType.equals(CellType.BOOLEAN)) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cellType.equals(CellType.NUMERIC)) {
            cellValue = String.valueOf(cell.getNumericCellValue());
        }

        return cellValue;
    }
}
