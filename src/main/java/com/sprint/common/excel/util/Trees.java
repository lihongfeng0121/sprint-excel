package com.sprint.common.excel.util;

import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * 树工具
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2021年01月12日
 */
public abstract class Trees {

    public static <T extends Tree<T>, K> List<T> castCollectionToTree(Collection<T> collection,
                                                                      Function<T, K> keyGetter, Function<T, K> parentKeyGetter) {

        Set<K> keySet = new HashSet<>();

        Map<K, List<T>> grouping = collection.stream().filter(item -> {
            keySet.add(keyGetter.apply(item));
            return parentKeyGetter.apply(item) != null;
        }).collect(Collectors.groupingBy(parentKeyGetter));

        List<T> roots = collection.stream().filter(item -> {
            K parentKey = parentKeyGetter.apply(item);
            return parentKey == null || Miscs.isEmpty(grouping.get(parentKey)) || !keySet.contains(parentKey);
        }).collect(Collectors.toList());

        List<T> rootList = new ArrayList<>();

        for (T root : roots) {
            fillTree(root, keyGetter, grouping);
            rootList.add(root);
        }

        return rootList;
    }

    private static <T extends Tree<T>, K> void fillTree(T root, Function<T, K> keyGetter, Map<K, List<T>> grouping) {
        List<T> childList = grouping.get(keyGetter.apply(root));

        if (Miscs.isEmpty(childList)) {
            return;
        }

        root.setChildren(new ArrayList<>());

        for (T child : childList) {
            child.setParent(root);
            root.getChildren().add(child);
            fillTree(child, keyGetter, grouping);
        }
    }
}