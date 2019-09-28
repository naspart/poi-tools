package com.naspat.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelTarget {
    String title() default "数据报表";

    String indexName() default "序号";

    int frozenColumns() default 0;

    int frozenRows() default 0;
}