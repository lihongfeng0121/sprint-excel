package com.sprint.common.excel.data;

import java.util.Objects;

public class ExcelCell implements Cloneable {

    public static final int NUMERIC_TYPE = 0;
    public static final int STRING_TYPE = 1;

    private Integer cellType;
    private Object value;
    private String comment;
    private boolean isCommentVisible = false;
    private ExcelCellStyle cellStyle;
    private int rowNumOffset = 0;
    private int columnNumOffset = 0;
    private int rowSize = 1;
    private int columnSize = 1;

    public Object getValue() {
        if (this.value == null) {
            this.value = "";
        }

        return this.value;
    }

    public ExcelCell setValue(Object value) {
        if (value == null) {
            value = "";
        }

        this.value = value;
        return this;
    }

    public ExcelCell() {
    }

    public ExcelCell(Object value) {
        this(value, 0);
    }

    public ExcelCell(Object value, int columnNumOffset) {
        this(value, null, columnNumOffset);
    }

    public ExcelCell(Object value, Integer cellType, int columnNumOffset) {
        this.value = value;
        this.cellType = cellType;
        this.columnNumOffset = columnNumOffset;
    }

    public static ExcelCell of(Object value) {
        return new ExcelCell(value);
    }

    public static ExcelCell of(Object value, int columnNumOffset) {
        return new ExcelCell(value, columnNumOffset);
    }

    public static ExcelCell of(Object value, Integer cellType, int columnNumOffset) {
        return new ExcelCell(value, cellType, columnNumOffset);
    }

    public Integer getCellType() {
        return this.cellType;
    }

    public String getComment() {
        return this.comment;
    }

    public boolean getIsCommentVisible() {
        return this.isCommentVisible;
    }

    public ExcelCellStyle getCellStyle() {
        return this.cellStyle;
    }

    public int getRowNumOffset() {
        return this.rowNumOffset;
    }

    public int getColumnNumOffset() {
        return this.columnNumOffset;
    }

    public int getRowSize() {
        return this.rowSize;
    }

    public int getColumnSize() {
        return this.columnSize;
    }

    public ExcelCell setCellType(Integer cellType) {
        this.cellType = cellType;
        return this;
    }

    public ExcelCell setComment(String comment) {
        this.comment = comment;
        return this;
    }

    public ExcelCell setIsCommentVisible(boolean isCommentVisible) {
        this.isCommentVisible = isCommentVisible;
        return this;
    }

    public ExcelCell setCellStyle(ExcelCellStyle cellStyle) {
        this.cellStyle = cellStyle;
        return this;
    }

    public ExcelCell setRowNumOffset(int rowNumOffset) {
        this.rowNumOffset = rowNumOffset;
        return this;
    }

    public ExcelCell setColumnNumOffset(int columnNumOffset) {
        this.columnNumOffset = columnNumOffset;
        return this;
    }

    public ExcelCell setRowSize(int rowSize) {
        this.rowSize = rowSize;
        return this;
    }

    public ExcelCell setColumnSize(int columnSize) {
        this.columnSize = columnSize;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        ExcelCell cell = (ExcelCell) o;
        return isCommentVisible == cell.isCommentVisible && rowNumOffset == cell.rowNumOffset
            && columnNumOffset == cell.columnNumOffset && rowSize == cell.rowSize && columnSize == cell.columnSize
            && Objects.equals(cellType, cell.cellType) && Objects.equals(value, cell.value)
            && Objects.equals(comment, cell.comment) && Objects.equals(cellStyle, cell.cellStyle);
    }

    @Override
    public int hashCode() {
        return Objects.hash(cellType, value, comment, isCommentVisible, cellStyle, rowNumOffset, columnNumOffset,
            rowSize, columnSize);
    }

    @Override
    public String toString() {
        return "ExcelCell(cellType=" + this.getCellType() + ", value=" + this.getValue() + ", comment="
            + this.getComment() + ", isCommentVisible=" + this.getIsCommentVisible() + ", cellStyle="
            + this.getCellStyle() + ", rowNumOffset=" + this.getRowNumOffset() + ", columnNumOffset="
            + this.getColumnNumOffset() + ", rowSize=" + this.getRowSize() + ", columnSize=" + this.getColumnSize()
            + ")";
    }

    @Override
    protected ExcelCell clone() {
        try {
            return (ExcelCell) super.clone();
        } catch (CloneNotSupportedException e) {
            e.printStackTrace();
        }
        return null;
    }
}
