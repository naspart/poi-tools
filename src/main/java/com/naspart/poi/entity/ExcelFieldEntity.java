package com.naspart.poi.entity;

import com.naspart.poi.enums.HorizontalAlignment;
import com.naspart.poi.enums.VerticalAlignment;
import lombok.Builder;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.Map;

@Data
@Builder
public class ExcelFieldEntity {
    private Field field;

    private String name;

    private String format;

    private int width;

    private HorizontalAlignment horizontalAlignment;

    private VerticalAlignment verticalAlignment;

    private Map<String, String> replaceMap;
}
