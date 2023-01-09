package com.sprint.common.excel.writer;

import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.writer.excelables.*;

import java.util.Map;
import java.util.function.BiFunction;

/**
 * 可导出的工具
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年02月18日
 */
public final class Excelables {

    public static final Excelable<?> NONE = new Excelable<Object>() {
        @Override
        public ExcelCell[] exportRowName() {
            return new ExcelCell[0];
        }

        @Override
        public ExcelCell[] exportRowValue(Object var1) {
            return new ExcelCell[0];
        }
    };

    public static <T> Excelable<T> none() {
        return (Excelable<T>) NONE;
    }


    public static <T extends ExcelExportable> Excelable<T> exportable(String[] headers) {
        return ExportableBeanSheetArray.of(headers);
    }


    public static <T> Excelable<T> of(Map<String, String> mapper) {
        return BeanSheetMapper.of(mapper);
    }

    public static <T> Excelable<T> ofNumber(Map<String, String> mapper) {
        return BeanSheetMapper.ofNumber(mapper);
    }

    public static <T> Excelable<T> of(String[] headers, String[] propertys) {
        return BeanSheetArray.of(headers, propertys);
    }

    public static <T> Excelable<T> ofNumber(String[] headers, String[] propertys) {
        return BeanSheetArray.ofNumber(headers, propertys);
    }

    public static <T> Excelable<T> of(String[] headers) {
        return BeanSheetArray.of(headers, headers);
    }


    public static <T> Excelable<T> of(String[] headers, String[] propertys, BiFunction<T, String, String> function) {
        return BeanSheetArrayFunction.of(headers, propertys, function);
    }

    public static <T> Excelable<T> of(Class<T> xCelBeanClass) {
        return XCellBeanSheet.of(xCelBeanClass, false);
    }
}