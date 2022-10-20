package com.sprint.common.excel.writer.excelables;

import com.sprint.common.converter.conversion.nested.bean.introspection.CachedIntrospectionResults;
import com.sprint.common.converter.conversion.nested.bean.introspection.PropertyAccess;
import com.sprint.common.converter.util.Types;
import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.data.XCell;
import com.sprint.common.excel.util.AbstractTree;
import com.sprint.common.excel.util.Excels;
import com.sprint.common.excel.util.Miscs;
import com.sprint.common.excel.util.Trees;
import com.sprint.common.excel.writer.Excelable;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Bean 注解可写Excel的
 *
 * @author hongfeng.li
 * @since 2022/7/18
 */
public class XCellBeanSheet<T> implements Excelable<T> {

    private static final ExcelCell[] CELL_ARRAY = new ExcelCell[0];

    private final Class<T> tClass;

    private final boolean numberCell;

    //
    static class FieldNode extends AbstractTree<FieldNode> {

        private final PropertyAccess field;

        private final PropertyAccess parentField;

        public FieldNode(PropertyAccess field, PropertyAccess parentField) {
            this.field = field;
            this.parentField = parentField;
        }

        public PropertyAccess getField() {
            return field;
        }

        public PropertyAccess getParentField() {
            return parentField;
        }
    }

    public XCellBeanSheet(Class<T> tClass, boolean numberCell) {
        this.tClass = tClass;
        this.numberCell = numberCell;
    }


    public static <T> XCellBeanSheet<T> of(Class<T> tClass, boolean numberCell) {
        return new XCellBeanSheet<>(tClass, numberCell);
    }

    @Override
    public ExcelCell[] exportRowName() {
        List<FieldNode> list = new ArrayList<>();
        fillRowTreeNode(list, null, getXCellProperties(tClass));
        List<FieldNode> fieldsNodeTree = Trees.castCollectionToTree(list, FieldNode::getField,
                FieldNode::getParentField);

        int maxDepth = 0;

        if (Miscs.isNotEmpty(fieldsNodeTree)) {
            for (FieldNode root : fieldsNodeTree) {
                maxDepth = Math.max(root.depth(), maxDepth);
            }
        }

        List<ExcelCell> excelCellList = new ArrayList<>();

        if (Miscs.isNotEmpty(fieldsNodeTree)) {
            int columnNumOffset = 0;
            for (FieldNode root : fieldsNodeTree) {
                fillRowName(excelCellList, maxDepth, columnNumOffset, 0, root);
                columnNumOffset += root.width();
            }
        }

        return excelCellList.toArray(CELL_ARRAY);
    }

    private void fillRowName(List<ExcelCell> excelCellList, int maxDepth, int columnNumOffset, int rowNumOffset,
                             FieldNode fieldNode) {
        ExcelCell excelCell = ExcelCell.of(((XCell) fieldNode.getField().getAnnotation(XCell.class)).name());
        excelCell.setColumnNumOffset(columnNumOffset);
        excelCell.setRowNumOffset(rowNumOffset);
        excelCell.setColumnSize(fieldNode.width());
        excelCellList.add(excelCell);
        int rowSize = maxDepth;

        if (Miscs.isNotEmpty(fieldNode.getChildren())) {
            int coffset = columnNumOffset;
            rowNumOffset++;
            int cmaxDepth = 0;
            for (FieldNode root : fieldNode.getChildren()) {
                cmaxDepth = Math.max(root.depth(), cmaxDepth);
            }

            rowSize -= cmaxDepth;

            for (FieldNode root : fieldNode.getChildren()) {
                fillRowName(excelCellList, cmaxDepth, coffset, rowNumOffset, root);
                coffset += root.width();
            }
        }
        excelCell.setRowSize(rowSize);
    }

    @Override
    public ExcelCell[] exportRowValue(T var1) {
        List<ExcelCell> list = new ArrayList<>();
        fillRowValue(list, var1, getXCellProperties(tClass));
        ExcelCell[] cells = list.toArray(CELL_ARRAY);
        Excels.setSimpleCellColumnOffset(cells);
        return cells;
    }

    private List<PropertyAccess> getXCellProperties(Class<?> tClass) {
        return Arrays.stream(CachedIntrospectionResults.forClass(tClass).getReadPropertyAccess()).filter(item -> item.getAnnotation(XCell.class) != null).collect(Collectors.toList());
    }

    private void fillRowTreeNode(List<FieldNode> nodes, PropertyAccess parentField, List<PropertyAccess> fieldList) {
        fieldList.sort(Comparator.comparingInt(field -> ((XCell) field.getAnnotation(XCell.class)).sort()));
        for (PropertyAccess field : fieldList) {
            XCell xCell = field.getAnnotation(XCell.class);
            // 如果需要展开
            if (Types.isBean(field.extractClass()) && xCell.titleUnfold()) {
                List<PropertyAccess> childFieldList = getXCellProperties(field.extractClass());
                if (Miscs.isNotEmpty(xCell.name())) {
                    nodes.add(new FieldNode(field, parentField));
                    if (Miscs.isNotEmpty(childFieldList)) {
                        fillRowTreeNode(nodes, field, getXCellProperties(field.extractClass()));
                    }
                } else {
                    if (Miscs.isNotEmpty(childFieldList)) {
                        fillRowTreeNode(nodes, field, getXCellProperties(field.extractClass()));
                    } else {
                        nodes.add(new FieldNode(field, parentField));
                    }
                }
            } else {
                nodes.add(new FieldNode(field, parentField));
            }
        }
    }

    private void fillRowValue(List<ExcelCell> values, Object obj, List<PropertyAccess> fieldList) {
        fieldList.sort(Comparator.comparingInt(field -> ((XCell) field.getAnnotation(XCell.class)).sort()));
        for (PropertyAccess field : fieldList) {
            XCell xCell = field.getAnnotation(XCell.class);
            Class<?> xCellClazz = field.extractClass();
            if (Types.isBean(xCellClazz) && xCell.titleUnfold()) {
                List<PropertyAccess> childFieldList = getXCellProperties(xCellClazz);
                if (Miscs.isNotEmpty(childFieldList)) {
                    try {
                        fillRowValue(values, obj == null ? null : field.getValue(obj), childFieldList);
                    } catch (Exception ignored) {
                        values.add(ExcelCell.of(""));
                    }
                } else {
                    try {
                        Object cellVal = obj == null ? "" : field.getValue(obj);
                        if (numberCell && cellVal instanceof Number) {
                            values.add(ExcelCell.of(cellVal, ExcelCell.NUMERIC_TYPE, 0));
                        } else {
                            values.add(ExcelCell.of(cellVal));
                        }
                    } catch (Exception ignored) {
                        values.add(ExcelCell.of(""));
                    }
                }
            } else if (xCell.dataUnfold() && Types.isCollection(field.extractClass())) {
                throw new UnsupportedOperationException();
            } else {
                try {
                    Object cellVal = obj == null ? "" : field.getValue(obj);
                    if (numberCell && cellVal instanceof Number) {
                        values.add(ExcelCell.of(cellVal, ExcelCell.NUMERIC_TYPE, 0));
                    } else {
                        values.add(ExcelCell.of(cellVal));
                    }
                } catch (Exception ignored) {
                    values.add(ExcelCell.of(""));
                }
            }
        }
    }
}
