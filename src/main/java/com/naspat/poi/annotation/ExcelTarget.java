package com.naspat.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelTarget {
    /**
     * @return sheet 标题
     */
    String title() default "数据报表";

    /**
     * @return 索引列名
     */
    String indexName() default "序号";

    /**
     * @return 冻结列
     */
    int frozenColumns() default 0;

    /**
     *  @return 冻结行
     */
    int frozenRows() default 0;
}