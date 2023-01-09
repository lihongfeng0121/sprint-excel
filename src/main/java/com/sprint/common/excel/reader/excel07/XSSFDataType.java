package com.sprint.common.excel.reader.excel07;

import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 *
 *
 * @author hongfeng.li
 * @since 2022/10/21
 */
public enum XSSFDataType {

    /**
     * boolean
     */
    BOOL("b"),
    /**
     * ex
     */
    ERROR("e"),
    /**
     * 公式
     */
    FORMULA("str"),
    /**
     * inlineStr
     */
    INLINESTR("inlineStr"),
    /**
     * s
     */
    SSTINDEX("s"),
    /**
     * d
     */
    NUMBER("d");


    private final String code;

    XSSFDataType(String code) {
        this.code = code;
    }

    public String getCode() {
        return code;
    }

    static final Map<String, XSSFDataType> CODE_MAP = Stream.of(XSSFDataType.values()).collect(Collectors.toMap(XSSFDataType::getCode, Function.identity()));

    public static XSSFDataType codeOf(String code) {
        return CODE_MAP.get(code);
    }

    public static XSSFDataType codeOf(String code, XSSFDataType defaultType) {
        return CODE_MAP.getOrDefault(code, defaultType);
    }
}
