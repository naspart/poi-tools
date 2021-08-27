package com.naspat.poi.entity;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

@Data
@Builder
public class ExcelTargetEntity {
    private boolean hasSheetTitle;

    private String sheetTitle;

    private int frozenColumns;

    private int frozenRows;

    private int mergeReferenceIndex;

    private List<Field> fields;

    private Map<Field, ExcelFieldEntity> excelFieldEntityMap;

    private Map<Field, CellStyle> cellStyleMap;
}
