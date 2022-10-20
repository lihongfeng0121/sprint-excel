package com.sprint.common.excel.util;

import java.util.Collection;
import java.util.Iterator;
import java.util.function.Supplier;

/**
 * @author hongfeng.li
 * @since 2022/10/20
 */
public class Safes {


    public static <T> T first(Collection<T> collection) {
        return at(collection, 0);
    }

    public static <T> T last(Collection<T> collection) {
        return at(collection, Miscs.size(collection) - 1);
    }

    public static <T> T at(Collection<T> collection, int i) {
        return at(collection, i, (T) null);
    }

    public static <T> T at(Collection<T> collection, int i, T defaultValue) {
        if (collection == null) {
            return null;
        } else {
            if (i < 0) {
                i += collection.size();
            }

            for (Iterator<T> iterator = collection.iterator(); iterator.hasNext(); --i) {
                T t = iterator.next();
                if (i == 0) {
                    return t;
                }
            }

            return defaultValue;
        }
    }


    public static <T> T ifNullThen(T value, T defaultValue) {
        return value == null ? defaultValue : value;
    }

    public static <T> T ifNullThen(T value, Supplier<T> supplier) {
        return value == null ? supplier.get() : value;
    }

    public static String ifEmptyThen(String value, String defaultValue) {
        return Miscs.isEmpty(value) ? defaultValue : value;
    }

    public static String ifEmptyThen(String value, Supplier<String> supplier) {
        return Miscs.isEmpty(value) ? supplier.get() : value;
    }
}
