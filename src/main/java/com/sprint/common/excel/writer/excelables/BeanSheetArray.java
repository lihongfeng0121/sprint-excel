package com.sprint.common.excel.writer.excelables;

import com.sprint.common.converter.conversion.nested.bean.Beans;
import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.writer.Excelable;

import java.util.ArrayList;
import java.util.List;

/**
 * bean可倒出的
 */
public class BeanSheetArray<T> implements Excelable<T> {

    private static final ExcelCell[] CELL_ARRAY = new ExcelCell[0];

    private final String[] headers;
    private final String[] propertys;

    private boolean numberCell;

    public BeanSheetArray(String[] headers, String[] propertys, boolean numberCell) {
        if (headers.length != propertys.length) {
            throw new IllegalArgumentException("headers/propertys 's length must be equal.");
        }
        this.headers = headers;
        this.propertys = propertys;
        this.numberCell = numberCell;
    }

    public BeanSheetArray(String[] headers, String[] propertys) {
        this(headers, propertys, false);
    }

    /**
     * @param headers   头数组
     * @param propertys 属性名
     * @param <T> t
     * @return ExcelCell
     */
    public static <T> BeanSheetArray<T> of(String[] headers, String[] propertys) {
        return new BeanSheetArray<>(headers, propertys);
    }

    public static <T> BeanSheetArray<T> ofNumber(String[] headers, String[] propertys) {
        return new BeanSheetArray<>(headers, propertys, true);
    }

    @Override
    public ExcelCell[] exportRowName() {
        List<ExcelCell> headerList = new ArrayList<>(headers.length);

        int columnNumOffset = 0;
        for (String header : headers) {
            headerList.add(ExcelCell.of(header, columnNumOffset++));
        }

        return headerList.toArray(CELL_ARRAY);
    }

    @Override
    public ExcelCell[] exportRowValue(T val) {
        if (val == null) {
            return CELL_ARRAY;
        }

        List<ExcelCell> columns = new ArrayList<>(propertys.length);

        int columnNumOffset = 0;
        for (String property : propertys) {
            Object cellVal = Beans.getProperty(val, property);
            if (numberCell && cellVal instanceof Number) {
                columns.add(ExcelCell.of(cellVal, ExcelCell.NUMERIC_TYPE, columnNumOffset++));
            } else {
                columns.add(ExcelCell.of(cellVal, columnNumOffset++));
            }
        }

        return columns.toArray(CELL_ARRAY);
    }
}
