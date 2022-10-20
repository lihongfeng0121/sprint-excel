package com.sprint.common.excel.data;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 标题数据
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年07月31日
 */
public class XTDRow extends XRow {

    private XRow titleRow;

    private Map<String, XCol> titleColMap;

    public XTDRow(XRow titleRow, XRow dataRow) {
        super(dataRow.getNumber(), dataRow.getCol());
        this.titleRow = titleRow;
        this.titleColMap = initTitleColMap();
    }

    private Map<String, XCol> initTitleColMap() {
        // 获取到遍历的当前行
        Map<String, XCol> map = new LinkedHashMap<>();
        List<XCol> cols = this.getCol();
        // 遍历当前行的列
        for (int j = 0; j < cols.size(); j++) {
            // map里存放的是列名称为key，该行该列为value
            if (titleRow.getCol().size() > j && titleRow.getCol().get(j) != null
                    && titleRow.getCol().get(j).getValue() != null
                    && titleRow.getCol().get(j).getValue().trim().length() > 0) {
                if (cols.size() >= j + 1 && cols.get(j) != null) {
                    map.put(titleRow.getCol().get(j).getValue().trim(), cols.get(j));
                } else {
                    map.put(titleRow.getCol().get(j).getValue().trim(), null);
                }
            }
        }
        return map;
    }

    public XRow getTitleRow() {
        return titleRow;
    }

    public Map<String, XCol> getTitleColMap() {
        return titleColMap;
    }
}