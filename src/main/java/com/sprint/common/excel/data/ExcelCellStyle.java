package com.sprint.common.excel.data;

import java.util.Objects;

import org.apache.poi.ss.usermodel.IndexedColors;

public class ExcelCellStyle {
    public static final short COLOR_BLACK;
    public static final short COLOR_WHITE;
    public static final short COLOR_GREY;
    public static final short COLOR_GREEN;
    public static final short COLOR_RED;
    public static final short COLOR_YELLOW;
    public static final short ALIGN_CENTER = 2;
    public static final short VERTICAL_CENTER = 1;
    public static final ExcelCellStyle HEAD_STYLE;
    private Short fontColor;
    private Short alignment;
    private Short verticalAlignment;
    private Boolean bold;

    public ExcelCellStyle() {
    }

    public Short getFontColor() {
        return this.fontColor;
    }

    public Short getAlignment() {
        return this.alignment;
    }

    public Short getVerticalAlignment() {
        return this.verticalAlignment;
    }

    public Boolean getBold() {
        return bold;
    }

    public void setFontColor(Short fontColor) {
        this.fontColor = fontColor;
    }

    public void setAlignment(Short alignment) {
        this.alignment = alignment;
    }

    public void setVerticalAlignment(Short verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public void setBold(Boolean bold) {
        this.bold = bold;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o)
            return true;
        if (o == null || getClass() != o.getClass())
            return false;
        ExcelCellStyle that = (ExcelCellStyle) o;
        return Objects.equals(fontColor, that.fontColor) && Objects.equals(alignment, that.alignment)
            && Objects.equals(verticalAlignment, that.verticalAlignment) && Objects.equals(bold, that.bold);
    }

    @Override
    public int hashCode() {
        return Objects.hash(fontColor, alignment, verticalAlignment, bold);
    }

    @Override
    public String toString() {
        return "ExcelCellStyle{" + "fontColor=" + fontColor + ", alignment=" + alignment + ", verticalAlignment="
            + verticalAlignment + ", bold=" + bold + '}';
    }

    static {
        COLOR_BLACK = IndexedColors.BLACK.index;
        COLOR_WHITE = IndexedColors.WHITE.index;
        COLOR_GREY = IndexedColors.GREY_80_PERCENT.index;
        COLOR_GREEN = IndexedColors.GREEN.index;
        COLOR_RED = IndexedColors.RED.index;
        COLOR_YELLOW = IndexedColors.YELLOW.index;
        HEAD_STYLE = new ExcelCellStyle();

        HEAD_STYLE.alignment = ALIGN_CENTER;
        HEAD_STYLE.verticalAlignment = VERTICAL_CENTER;
    }
}
