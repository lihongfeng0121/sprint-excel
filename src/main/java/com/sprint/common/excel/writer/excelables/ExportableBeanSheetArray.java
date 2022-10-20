package com.sprint.common.excel.writer.excelables;

import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.writer.Excelable;

/**
 * @author hongfeng.li
 * @since 2022/7/18
 */
public class ExportableBeanSheetArray<T extends ExcelExportable> implements Excelable<T> {

    private final ExcelCell[] headers;

    public static <T extends ExcelExportable> ExportableBeanSheetArray<T> of(String[] headers) {
        return new ExportableBeanSheetArray<>(headers);
    }

    public ExportableBeanSheetArray(String[] headerStrs) {
        ExcelCell[] headers = new ExcelCell[headerStrs.length];
        for (int i = 0; i < headerStrs.length; i++) {
            headers[i] = ExcelCell.of(headerStrs[i], i);
        }
        this.headers = headers;
    }

    public ExportableBeanSheetArray(ExcelCell[] headers) {
        this.headers = headers;
    }

    @Override
    public ExcelCell[] exportRowName() {
        return headers;
    }

    @Override
    public ExcelCell[] exportRowValue(ExcelExportable rowDate) {
        return rowDate.rowValue();
    }
}
