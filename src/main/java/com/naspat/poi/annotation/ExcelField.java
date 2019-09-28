package com.naspat.poi.annotation;

import com.naspat.poi.enums.HorizontalAlignment;
import com.naspat.poi.enums.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelField {
    /**
     * 字段名
     */
    String name();

    /**
     * 格式
     */
    String format() default "";

    String[] replace() default {};

    /**
     * 长度
     */
    int width() default 10;

    /**
     * 水平对齐
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.GENERAL;

    /**
     * 垂直对齐
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;
}
