package com.sprint.common.excel.data;

import java.util.List;

/**
 * SheetXRow
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2021年04月27日
 */
public class Sheet {

    private int sheetNumber;
    private String sheetName;
    private List<XRow> rows;

    public Sheet(int sheetNumber, String sheetName, List<XRow> rows) {
        this.sheetNumber = sheetNumber;
        this.sheetName = sheetName;
        this.rows = rows;
    }

    public int getSheetNumber() {
        return sheetNumber;
    }

    public void setSheetNumber(int sheetNumber) {
        this.sheetNumber = sheetNumber;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<XRow> getRows() {
        return rows;
    }

    public void setRows(List<XRow> rows) {
        this.rows = rows;
    }
}