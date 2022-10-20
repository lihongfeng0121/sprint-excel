package com.sprint.common.excel.util;

import java.beans.Transient;
import java.io.Serializable;
import java.util.Collection;

/**
 * 树
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2021年01月12日
 */
public interface Tree<T extends Tree<T>> extends Cloneable, Serializable {

    /**
     * 树层级
     *
     * @return
     */
    default int level() {
        int level = 0;

        Tree<T> parent = getParent();

        if (parent == null) {
            return level;
        }

        do {
            level++;
        } while ((parent = parent.getParent()) != null);

        return level;
    }

    /**
     * 树深度
     *
     * @return
     */
    default int depth() {
        if (getChildren() == null || getChildren().isEmpty()) {
            return 1;
        }

        int depth = 1;

        for (Tree<T> c : getChildren()) {
            if (c == null) {
                continue;
            }
            depth = Math.max(c.depth(), depth);
        }

        return depth + 1;
    }

    default int width() {
        if (getChildren() == null || getChildren().isEmpty() || getChildren().size() == 1) {
            return 1;
        }

        int width = getChildren().size();

        for (Tree<T> c : getChildren()) {
            if (c == null) {
                continue;
            }

            if (c.width() > 1) {
                width += c.width() - 1;
            }
        }

        return width;
    }

    /**
     * 父节点
     *
     * @return
     */
    @Transient
    T getParent();

    /**
     * 设置父节点
     *
     * @param parent
     * @return
     */
    void setParent(T parent);

    /**
     * 获取子节点
     *
     * @return
     */
    Collection<T> getChildren();

    /**
     * 设置子节点
     *
     * @param children
     */
    void setChildren(Collection<T> children);
}