package com.naspat.poi.function;

import java.util.Collection;

/**
 * 分页查询
 */
public interface ExcelDataFunction<P, T> {
    /**
     * 分页查询方法
     */
    Collection<T> getData(P param, int pageNum, int pageSize);

    /**
     * 集合内对象转换
     */
    T convert(T queryResult);
}
