package com.sprint.common.excel.data;

import com.sprint.common.converter.util.AbstractValue;

public class XCol extends AbstractValue {

    public static final XCol EMPTY = new XCol("");

    private final String value;

    public XCol(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }

    public boolean isEmpty() {
        int strLen;
        if (value != null && (strLen = value.length()) != 0) {
            for (int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(value.charAt(i))) {
                    return false;
                }
            }
        }
        return true;
    }

    @Override
    public String toString() {
        return value;
    }
}
