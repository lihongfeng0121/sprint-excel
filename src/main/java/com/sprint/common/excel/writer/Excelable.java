package com.sprint.common.excel.writer;

import com.sprint.common.excel.data.ExcelCell;

/**
 * 可导出的
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2019年10月17日
 */
public interface Excelable<T> {

    /**
     * 标题
     *
     * @return 标题行数据
     */
    ExcelCell[] exportRowName();

    /**
     * 行数据
     *
     * @param obj 行数据BEAN
     * @return 行数据
     */
    ExcelCell[] exportRowValue(T obj);

}