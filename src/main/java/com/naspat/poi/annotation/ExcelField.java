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
     * @return 字段名
     */
    String name();

    /**
     * @return 格式
     */
    String format() default "";

    String[] replace() default {};

    /**
     * @return 长度
     */
    int width() default 10;

    /**
     * @return 水平对齐
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.GENERAL;

    /**
     * @return 垂直对齐
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

    /**
     * @return 纵向合并内容相同的单元格
     */
    boolean verticalMerge() default false;
}
