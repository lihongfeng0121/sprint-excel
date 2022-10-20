package com.sprint.common.excel.data;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public enum ExcelType {

    /**
     * 2003
     */
    XLS(1, ".xls", "application/vnd.ms-excel", 65536),
    /**
     * 2007
     */
    XLSX(2, ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 1048576);

    private int id;
    private String suffix;
    private String contentType;
    private int maxRow;

    ExcelType(int id, String suffix, String contentType, int maxRow) {
        this.id = id;
        this.suffix = suffix;
        this.contentType = contentType;
        this.maxRow = maxRow;
    }

    public static ExcelType fromWorkbook(Workbook wb) {
        if (wb == null) {
            return null;
        } else {
            return wb instanceof HSSFWorkbook ? XLS : XLSX;
        }
    }

    public int getId() {
        return this.id;
    }

    public String getSuffix() {
        return this.suffix;
    }

    public String getContentType() {
        return this.contentType;
    }

    public int getMaxRow() {
        return this.maxRow;
    }
}
