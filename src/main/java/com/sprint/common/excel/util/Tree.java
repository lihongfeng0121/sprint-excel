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


    @Transient
    T getParent();


    void setParent(T parent);


    Collection<T> getChildren();


    void setChildren(Collection<T> children);
}