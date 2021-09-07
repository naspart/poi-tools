package com.hgobox.poi.entity;

import com.hgobox.poi.enums.VerticalAlignment;
import com.hgobox.poi.enums.HorizontalAlignment;
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

    private HorizontalAlignment horizontalAlignment;

    private VerticalAlignment verticalAlignment;

    private Map<String, String> replaceMap;
}
