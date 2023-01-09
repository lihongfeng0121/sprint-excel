package com.sprint.common.excel.writer.excelables;

import com.sprint.common.excel.data.ExcelCell;

/**
 * 可倒出的
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年08月05日
 */
public interface ExcelExportable {

    /**
     * 行数据
     *
     * @return 行数据
     */
    ExcelCell[] rowValue();
}