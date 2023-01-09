package com.sprint.common.excel.util;

import org.apache.poi.EmptyFileException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Objects;

/**
 * 文件工具类
 *
 * @since 2018-8-14
 */
public class FileUtils {

    private static final Logger logger = LoggerFactory.getLogger(FileUtils.class);

    public static boolean isExcel(InputStream inputStream) {
        try {
            FileType fileType = FileType.valueOf(inputStream);
            if (Objects.equals(fileType, FileType.OLE2) || Objects.equals(fileType, FileType.OOXML)) {
                return true;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    public static boolean createFolder(String folder) {
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

    public static String resolvePath(String... pathMore) {
        StringBuilder path = new StringBuilder();
        for (String pstr : pathMore) {
            if (path.lastIndexOf(File.separator) != (path.length() - File.separator.length())) {
                path.append(File.separator);
            }
            path.append(pstr);
        }

        return path.toString();
    }

    private static int BYTE_ARRAY_MAX_OVERRIDE = -1;

    private static void checkByteSizeLimit(int length) throws IOException {
        if (BYTE_ARRAY_MAX_OVERRIDE != -1 && length > BYTE_ARRAY_MAX_OVERRIDE) {
            throwRFE(length, BYTE_ARRAY_MAX_OVERRIDE);
        }
    }

    public static void write(InputStream in, OutputStream outputStream) throws IOException {
        byte[] buffer = new byte[512];

        int lg;
        while ((lg = in.read(buffer)) > 0) {
            outputStream.write(buffer, 0, lg);
        }
    }

    public static byte[] peekFirstNBytes(InputStream stream, int limit) throws IOException {
        checkByteSizeLimit(limit);
        stream.mark(limit);
        ByteArrayOutputStream bos = new ByteArrayOutputStream(limit);
        write(new BoundedInputStream(stream, limit), bos);

        int readBytes = bos.size();
        if (readBytes == 0) {
            throw new EmptyFileException();
        }

        if (readBytes < limit) {
            bos.write(new byte[limit - readBytes]);
        }
        byte[] peekedBytes = bos.toByteArray();
        if (stream instanceof PushbackInputStream) {
            PushbackInputStream pin = (PushbackInputStream) stream;
            pin.unread(peekedBytes, 0, readBytes);
        } else {
            stream.reset();
        }

        return peekedBytes;
    }


    private static void throwRFE(long length, int maxLength) throws IOException {
        throw new IOException("Tried to allocate an array of length " + length + ", but " + maxLength
                + " is the maximum for this record type.\n"
                + "If the file is not corrupt, please open an issue on bugzilla to request \n"
                + "increasing the maximum allowable size for this record type.\n"
                + "As a temporary workaround, consider setting a higher override value with "
                + "IOUtils.setByteArrayMaxOverride()");
    }
}
