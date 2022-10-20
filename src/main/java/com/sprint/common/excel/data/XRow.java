package com.sprint.common.excel.data;

import java.util.ArrayList;
import java.util.List;

public class XRow {

    private Integer number;

    private List<XCol> col;

    public XRow(Integer number) {
        this.number = number;
        col = new ArrayList<>();
    }

    public XRow(Integer number, List<XCol> col) {
        this.number = number;
        this.col = col;
    }

    public List<XCol> getCol() {
        return col;
    }

    public Integer getNumber() {
        return number;
    }

    public boolean isEmpty() {
        for (XCol c : col) {
            if (!c.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    @Override
    public String toString() {
        return "XRow{" + "number=" + number + ", col=" + col + '}';
    }
}
