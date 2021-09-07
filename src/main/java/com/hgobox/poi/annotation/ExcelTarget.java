package com.hgobox.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation for mapping excel sheet attributes.
 *
 * <p><strong>Note:</strong> This annotation must be used at the class level.
 *
 * @author Fabron Lau
 * @since 1.0.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelTarget {
    /**
     * @return sheet 标题
     */
    String title() default "";

    /**
     * @return 冻结列
     */
    int frozenColumns() default 0;

    /**
     * @return 冻结行
     */
    int frozenRows() default 0;
}