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

    ExcelCell[] exportRowName();

    ExcelCell[] exportRowValue(T var1);

}