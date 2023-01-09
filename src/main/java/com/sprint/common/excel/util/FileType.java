package com.sprint.common.excel.util;


import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;

import static com.sprint.common.excel.util.FileHeaders.OOXML_FILE_HEADER;
import static com.sprint.common.excel.util.FileHeaders.RAW_XML_FILE_HEADER;


/**
 * 文件类型
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2021年02月08日
 */
public enum FileType {

    /**
     * OLE2 / BIFF8+ stream used for Office 97 and higher documents
     */
    OLE2(0xE11AB1A1E011CFD0L),
    /**
     * OOXML / ZIP stream
     */
    OOXML(OOXML_FILE_HEADER),
    /**
     * XML file
     */
    XML(RAW_XML_FILE_HEADER),
    /**
     * BIFF2 raw stream - for Excel 2
     */
    BIFF2(new byte[]{0x09, 0x00, // sid=0x0009
            0x04, 0x00, // size=0x0004
            0x00, 0x00, // unused
            '?', 0x00 // '?' = multiple values
    }),
    /**
     * BIFF3 raw stream - for Excel 3
     */
    BIFF3(new byte[]{0x09, 0x02, // sid=0x0209
            0x06, 0x00, // size=0x0006
            0x00, 0x00, // unused
            '?', 0x00 // '?' = multiple values
    }),
    /**
     * BIFF4 raw stream - for Excel 4
     */
    BIFF4(new byte[]{0x09, 0x04, // sid=0x0409
            0x06, 0x00, // size=0x0006
            0x00, 0x00, // unused
            '?', 0x00 // '? = multiple values
    }, new byte[]{0x09, 0x04, // sid=0x0409
            0x06, 0x00, // size=0x0006
            0x00, 0x00, // unused
            0x00, 0x01}),
    /**
     * Old MS Write raw stream
     */
    MSWRITE(new byte[]{0x31, (byte) 0xbe, 0x00, 0x00}, new byte[]{0x32, (byte) 0xbe, 0x00, 0x00}),
    /**
     * RTF document
     */
    RTF("{\\rtf"),
    /**
     * PDF document
     */
    PDF("%PDF"),
    /**
     * Some different HTML documents
     */
    HTML("<!DOCTYP", "<html", "\n\r<html", "\r\n<html", "\r<html", "\n<html", "<HTML", "\r\n<HTML", "\n\r<HTML",
            "\r<HTML", "\n<HTML"), WORD2(new byte[]{(byte) 0xdb, (byte) 0xa5, 0x2d, 0x00}),
    /**
     * JPEG image
     */
    JPEG(new byte[]{(byte) 0xFF, (byte) 0xD8, (byte) 0xFF, (byte) 0xDB},
            new byte[]{(byte) 0xFF, (byte) 0xD8, (byte) 0xFF, (byte) 0xE0, '?', '?', 'J', 'F', 'I', 'F', 0x00, 0x01},
            new byte[]{(byte) 0xFF, (byte) 0xD8, (byte) 0xFF, (byte) 0xEE},
            new byte[]{(byte) 0xFF, (byte) 0xD8, (byte) 0xFF, (byte) 0xE1, '?', '?', 'E', 'x', 'i', 'f', 0x00, 0x00}),
    /**
     * GIF image
     */
    GIF("GIF87a", "GIF89a"),
    /**
     * PNG Image
     */
    PNG(new byte[]{(byte) 0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A}),
    /**
     * TIFF Image
     */
    TIFF("II*\u0000", "MM\u0000*"),
    /**
     * WMF image with a placeable header
     */
    WMF(new byte[]{(byte) 0xD7, (byte) 0xCD, (byte) 0xC6, (byte) 0x9A}),
    /**
     * EMF image
     */
    EMF(new byte[]{1, 0, 0, 0, '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?',
            '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', '?', ' ', 'E', 'M',
            'F'}),
    /**
     * BMP image
     */
    BMP(new byte[]{'B', 'M'}),
    // keep UNKNOWN always as last enum!
    /**
     * UNKNOWN type
     */
    UNKNOWN(new byte[0]);

    /**
     * update this if a longer pattern is added
     */
    final static int MAX_PATTERN_LENGTH = 44;

    final byte[][] type;

    FileType(long type) {
        this.type = new byte[1][8];
        Miscs.putLong(this.type[0], 0, type);
    }

    FileType(byte[]... type) {
        this.type = type;
    }

    FileType(String... type) {
        this.type = new byte[type.length][];
        int i = 0;
        for (String s : type) {
            this.type[i++] = s.getBytes(Charset.forName("CP1252"));
        }
    }

    public static FileType valueOf(byte[] type) {
        for (FileType fm : values()) {
            for (byte[] ma : fm.type) {
                // don't try to match if the given byte-array is too short
                // for this pattern anyway
                if (type.length < ma.length) {
                    continue;
                }

                if (findType(ma, type)) {
                    return fm;
                }
            }
        }
        return UNKNOWN;
    }

    private static boolean findType(byte[] expected, byte[] actual) {
        int i = 0;
        for (byte expectedByte : expected) {
            if (actual[i++] != expectedByte && expectedByte != '?') {
                return false;
            }
        }
        return true;
    }

    public static FileType valueOf(InputStream inp) throws IOException {
        if (!inp.markSupported()) {
            inp = prepareToCheckType(inp);
        }

        // Grab the first bytes of this stream
        byte[] data = FileUtils.peekFirstNBytes(inp, MAX_PATTERN_LENGTH);

        return FileType.valueOf(data);
    }

    public static InputStream prepareToCheckType(InputStream stream) {
        if (stream.markSupported()) {
            return stream;
        }
        // we used to process the data via a PushbackInputStream, but user code could provide a too small one
        // so we use a BufferedInputStream instead now
        return new BufferedInputStream(stream);
    }
}