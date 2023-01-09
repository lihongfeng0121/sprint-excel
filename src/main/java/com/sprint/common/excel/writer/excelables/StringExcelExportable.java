package com.sprint.common.excel.writer.excelables;

import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.util.Excels;

import java.util.Optional;
import java.util.stream.Stream;

/**
 * 字符类型可倒出的
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年08月05日
 */
public abstract class StringExcelExportable implements ExcelExportable {

    @Override
    public ExcelCell[] rowValue() {
        ExcelCell[] rowValues = Stream.of(Optional.ofNullable(stringRowValue()).orElse(new String[0]))
                .map(ExcelCell::of).toArray(ExcelCell[]::new);
        Excels.setSimpleCellColumnOffset(rowValues);
        return rowValues;
    }

    /**
     * 字符串行数据
     *
     * @return 行数据
     */
    public abstract String[] stringRowValue();
}