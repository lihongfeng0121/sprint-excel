package com.sprint.common.excel.util;

import java.beans.Transient;
import java.util.Collection;
import java.util.Collections;

/**
 * 树
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年02月05日
 */
public abstract class AbstractTree<T extends AbstractTree<T>> implements Tree<T> {

    private static final long serialVersionUID = 3312459054550526627L;

    private transient T parent;

    private Collection<T> children = Collections.emptyList();

    @Override
    @Transient
    public T getParent() {
        return parent;
    }

    @Override
    public void setParent(T parent) {
        this.parent = parent;
    }

    @Override
    public Collection<T> getChildren() {
        return children;
    }

    @Override
    public void setChildren(Collection<T> children) {
        this.children = children;
    }
}