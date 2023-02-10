package com.sprint.common.excel.writer.excelables;

/**
 * @author hongfeng.li
 * @since 2022/7/18
 */

import com.sprint.common.converter.util.Beans;
import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.writer.Excelable;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 头/属性映射
 */
public class BeanSheetMapper<T> implements Excelable<T> {

    private static final ExcelCell[] CELL_ARRAY = new ExcelCell[0];

    private Map<String, String> mapper;

    private boolean numberCell;

    public BeanSheetMapper(Map<String, String> mapper) {
        this(mapper, false);
    }

    public BeanSheetMapper(Map<String, String> mapper, boolean numberCell) {
        if (mapper == null) {
            throw new IllegalArgumentException("mapper can't be null!");
        }
        this.mapper = mapper;
        this.numberCell = numberCell;
    }

    public static <T> BeanSheetMapper<T> of(Map<String, String> mapper) {
        return new BeanSheetMapper<>(mapper);
    }

    public static <T> BeanSheetMapper<T> ofNumber(Map<String, String> mapper) {
        return new BeanSheetMapper<>(mapper, true);
    }

    @Override
    public ExcelCell[] exportRowName() {
        List<ExcelCell> headers = new ArrayList<>(mapper.size());

        int columnNumOffset = 0;
        for (Map.Entry<String, String> entry : mapper.entrySet()) {
            headers.add(ExcelCell.of(entry.getKey(), columnNumOffset++));
        }

        return headers.toArray(CELL_ARRAY);
    }

    @Override
    public ExcelCell[] exportRowValue(T val) {
        if (val == null) {
            return CELL_ARRAY;
        }

        List<ExcelCell> columns = new ArrayList<>(mapper.size());

        int columnNumOffset = 0;
        for (Map.Entry<String, String> entry : mapper.entrySet()) {
            Object cellVal = Beans.getProperty(val, entry.getValue());
            if (numberCell && cellVal instanceof Number) {
                columns.add(ExcelCell.of(cellVal, ExcelCell.NUMERIC_TYPE, columnNumOffset++));
            } else {
                columns.add(ExcelCell.of(cellVal, columnNumOffset++));
            }
        }

        return columns.toArray(CELL_ARRAY);
    }
}
