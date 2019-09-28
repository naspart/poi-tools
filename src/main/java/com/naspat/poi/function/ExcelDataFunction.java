package com.naspat.poi.function;

import java.util.Collection;

public interface ExcelDataFunction<P, T> {
    Collection<T> getData(P param, int pageNum, int pageSize);

    T convert(T queryResult);
}
