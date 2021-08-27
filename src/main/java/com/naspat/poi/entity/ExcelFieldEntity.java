package com.naspat.poi.entity;

import com.naspat.poi.enums.HorizontalAlignment;
import com.naspat.poi.enums.VerticalAlignment;
import lombok.Builder;
import lombok.Data;
import lombok.Getter;

import java.lang.reflect.Field;
import java.util.Map;

@Getter
@Data
@Builder
public class ExcelFieldEntity {
    private Field field;

    private String name;

    private String format;

    private int width;

    private boolean verticalMerge;

    private HorizontalAlignment horizontalAlignment;

    private VerticalAlignment verticalAlignment;

    private Map<String, String> replaceMap;
}
