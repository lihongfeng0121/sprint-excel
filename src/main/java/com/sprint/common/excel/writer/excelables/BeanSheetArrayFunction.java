package com.sprint.common.excel.writer.excelables;

import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.writer.Excelable;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiFunction;

/**
 * bean 属性解析成 ExcelCell
 *
 * @author hongfeng.li
 * @since 2022/7/18
 */
public class BeanSheetArrayFunction<T> implements Excelable<T> {

    private static final ExcelCell[] CELL_ARRAY = new ExcelCell[0];

    private final String[] headers;
    private final String[] propertys;
    private final BiFunction<T, String, String> function;

    public BeanSheetArrayFunction(String[] headers, String[] propertys, BiFunction<T, String, String> function) {
        this.headers = headers;
        this.propertys = propertys;
        this.function = function;
    }


    public static <T> BeanSheetArrayFunction<T> of(String[] headers, String[] propertys,
                                                   BiFunction<T, String, String> function) {
        return new BeanSheetArrayFunction<>(headers, propertys, function);
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
            columns.add(ExcelCell.of(function.apply(val, property), columnNumOffset++));
        }

        return columns.toArray(CELL_ARRAY);
    }
}
