package com.sprint.common.excel.util;

import com.sprint.common.converter.conversion.nested.bean.Beans;
import com.sprint.common.converter.util.Types;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;

/**
 * @author hongfeng.li
 * @since 2022/10/8
 */
public class Miscs {

    private static final Logger logger = LoggerFactory.getLogger(Miscs.class);

    /**
     * 是否为空串
     *
     * @param cs 字符
     * @return boolean
     */
    public static boolean isBlank(CharSequence cs) {
        int strLen;
        if (cs != null && (strLen = cs.length()) != 0) {
            for (int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(cs.charAt(i))) {
                    return false;
                }
            }
            return true;
        } else {
            return true;
        }
    }

    /**
     * 是否为空串
     *
     * @param cs 字符
     * @return boolean
     */
    public static boolean isNotBlank(CharSequence cs) {
        return !isBlank(cs);
    }

    /**
     * 是否为空
     *
     * @param cs 字符
     * @return boolean
     */
    public static boolean isEmpty(CharSequence cs) {
        return cs == null || cs.length() == 0;
    }

    /**
     * 是否不为空
     *
     * @param cs 字符
     * @return boolean
     */
    public static boolean isNotEmpty(CharSequence cs) {
        return !isEmpty(cs);
    }

    /**
     * 集合是否为空
     *
     * @param collection 集合
     * @return 返回是否为空
     */
    public static boolean isEmpty(Collection<?> collection) {
        return !isNotEmpty(collection);
    }

    /**
     * 集合是否不为空
     *
     * @param collection 集合
     * @return 返回是否不为空
     */
    public static boolean isNotEmpty(Collection<?> collection) {
        return collection != null && collection.size() > 0;
    }

    /**
     * 是否不为空
     *
     * @param map map
     * @return 返回是否不为空
     */
    public static boolean isNotEmpty(Map map) {
        return map != null && !map.isEmpty();
    }

    /**
     * 是否为空
     *
     * @param map map
     * @return 返回是否为空
     */
    public static boolean isEmpty(Map map) {
        return !isNotEmpty(map);
    }


    /**
     * 重写map的key，val
     *
     * @param map       源map
     * @param keyWriter key重写方法
     * @param valWriter val重写方法
     * @param <K>       源Map key的值类型
     * @param <V>       源Map val的值类型
     * @param <K2>      目标Map的key值类型
     * @param <V2>      目标Map的val值类型
     * @return 返回重写k/v后的Map
     */
    public static <K, V, K2, V2> Map<K2, V2> rewrite(Map<K, V> map, Function<K, K2> keyWriter,
                                                     Function<V, V2> valWriter) {

        Map<K2, V2> rewrite;

        if (Types.getConstructorIfAvailable(map.getClass()) != null) {
            rewrite = Beans.instanceMap(map.getClass());
        } else if (map instanceof HashMap) {
            rewrite = new HashMap<>(map.size());
        } else if (map instanceof ConcurrentHashMap) {
            rewrite = new ConcurrentHashMap<>(map.size());
        } else {
            rewrite = new LinkedHashMap<>(map.size());
        }

        map.forEach((key, val) -> rewrite.put(keyWriter.apply(key), valWriter.apply(val)));

        return rewrite;
    }

    /**
     * 重写Map的值
     *
     * @param map       源map
     * @param keyWriter key重写方法
     * @param <K>       源Map key的范型
     * @param <K2>      目标Map的key值
     * @param <V>       源Map val的范型
     * @return
     */
    public static <K, V, K2> Map<K2, V> rewriteKey(Map<K, V> map, Function<K, K2> keyWriter) {
        Map<K2, V> rewrite;
        if (Types.getConstructorIfAvailable(map.getClass()) != null) {
            rewrite = Beans.instanceMap(map.getClass());
        } else if (map instanceof HashMap) {
            rewrite = new HashMap<>(map.size());
        } else if (map instanceof ConcurrentHashMap) {
            rewrite = new ConcurrentHashMap<>(map.size());
        } else {
            rewrite = new LinkedHashMap<>(map.size());
        }

        map.forEach((key, val) -> rewrite.put(keyWriter.apply(key), val));

        return rewrite;
    }

    /**
     * 重写Map的值
     *
     * @param map       源map
     * @param valWriter val重写方法
     * @param <K>       源Map key的范型
     * @param <V>       源Map val的范型
     * @param <V2>      目标Map的Val值
     * @return 返回重写Map val后的map
     */
    public static <K, V, V2> Map<K, V2> rewriteVal(Map<K, V> map, Function<V, V2> valWriter) {
        Map<K, V2> rewrite;
        if (Types.getConstructorIfAvailable(map.getClass()) != null) {
            rewrite = Beans.instanceMap(map.getClass());
        } else if (map instanceof HashMap) {
            rewrite = new HashMap<>(map.size());
        } else if (map instanceof ConcurrentHashMap) {
            rewrite = new ConcurrentHashMap<>(map.size());
        } else {
            rewrite = new LinkedHashMap<>(map.size());
        }

        map.forEach((key, val) -> rewrite.put(key, valWriter.apply(val)));

        return rewrite;
    }


    /**
     * 翻转K-V
     *
     * @param map 源map
     * @return 返回翻转后的map
     */
    public static <K, V> Map<V, K> reverse(Map<K, V> map) {
        Map<V, K> reversed;
        if (Types.getConstructorIfAvailable(map.getClass()) != null) {
            reversed = Beans.instanceMap(map.getClass());
        } else if (map instanceof HashMap) {
            reversed = new HashMap<>(map.size());
        } else if (map instanceof ConcurrentHashMap) {
            reversed = new ConcurrentHashMap<>(map.size());
        } else {
            reversed = new LinkedHashMap<>(map.size());
        }

        map.forEach((key, val) -> {
            reversed.put(val, key);
        });
        return reversed;
    }


    /**
     * 集合大小
     *
     * @param collection 集合
     * @return
     */
    public static int size(Collection<?> collection) {
        if (collection == null) {
            return 0;
        }

        return collection.size();
    }

    /**
     * put a long value into a byte array
     *
     * @param data   the byte array
     * @param offset a starting offset into the byte array
     * @param value  the long (64-bit) value
     */
    public static void putLong(byte[] data, int offset, long value) {
        data[offset + 0] = (byte) ((value >>> 0) & 0xFF);
        data[offset + 1] = (byte) ((value >>> 8) & 0xFF);
        data[offset + 2] = (byte) ((value >>> 16) & 0xFF);
        data[offset + 3] = (byte) ((value >>> 24) & 0xFF);
        data[offset + 4] = (byte) ((value >>> 32) & 0xFF);
        data[offset + 5] = (byte) ((value >>> 40) & 0xFF);
        data[offset + 6] = (byte) ((value >>> 48) & 0xFF);
        data[offset + 7] = (byte) ((value >>> 56) & 0xFF);
    }

    /**
     * @param input
     * @param size
     * @return
     * @throws IOException
     */
    public static InputStream[] copyInputStream(InputStream input, int size) throws IOException {
        // 将InputStream对象转换成ByteArrayOutputStream
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        while ((len = input.read(buffer)) > -1) {
            byteArrayOutputStream.write(buffer, 0, len);
        }
        byteArrayOutputStream.flush();

        InputStream[] inputStreams = new InputStream[size];

        for (int i = 0; i < size; i++) {
            inputStreams[i] = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
        }

        return inputStreams;
    }

    public static String resolveFilePath(String... pathMore) {
        StringBuilder path = new StringBuilder();
        for (String pstr : pathMore) {
            if (path.lastIndexOf(File.separator) != (path.length() - File.separator.length())) {
                path.append(File.separator);
            }
            path.append(pstr);
        }

        return path.toString();
    }


    /**
     * 创建文件夹
     *
     * @param folder
     * @return
     */
    public static boolean createFileFolder(String folder) {
        File folderPath = new File(folder);
        try {
            if (!folderPath.exists()) {
                return folderPath.mkdir();
            }
        } catch (Exception e) {
            logger.error("[FileUtil][saveFile]mkdir error!", e);
            return false;
        }
        return true;
    }
}
