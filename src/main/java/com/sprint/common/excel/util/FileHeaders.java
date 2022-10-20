package com.sprint.common.excel.util;

/**
 * 文件头
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2021年02月08日
 */
public class FileHeaders {

    /**
     * The first 4 bytes of an OOXML file, used in detection
     */
    public static final byte[] OOXML_FILE_HEADER = new byte[]{0x50, 0x4b, 0x03, 0x04};
    /**
     * The first 5 bytes of a raw XML file, used in detection
     */
    public static final byte[] RAW_XML_FILE_HEADER = new byte[]{0x3c, 0x3f, 0x78, 0x6d, 0x6c};
}