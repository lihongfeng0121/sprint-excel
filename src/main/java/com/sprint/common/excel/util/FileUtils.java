package com.sprint.common.excel.util;

import org.apache.poi.EmptyFileException;
import org.apache.poi.util.BoundedInputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
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

    public static void writeBytesFile(byte[] bs, String filePath) throws Exception {
        File file = new File(filePath);
        File parentFile = file.getParentFile();
        if (!parentFile.exists()) {
            if (!parentFile.mkdirs()) {
                throw new IllegalStateException("create dir fail");
            }
        }
        try (FileOutputStream fos = new FileOutputStream(file)) {
            fos.write(bs);
            fos.flush();
            logger.info("write file path:{} length:{}", file.getCanonicalPath(), bs.length);
        }
    }


    public static void writeBytesFile(InputStream in, String filePath) throws Exception {
        File file = new File(filePath);
        File parentFile = file.getParentFile();
        if (!parentFile.exists()) {
            if (!parentFile.mkdirs()) {
                throw new IllegalStateException("create dir fail");
            }
        }

        FileOutputStream fos = new FileOutputStream(file);
        try {

            byte[] buffer = new byte[512];

            int len = 0;
            while ((len = in.read(buffer)) > 0) {
                fos.write(buffer, 0, len);
            }

            fos.flush();
            LoggerFactory.getLogger(FileUtils.class).info("write file path:{}", file.getCanonicalPath());
        } finally {
            fos.close();
        }
    }

    public static byte[] readBytesFile(String filePath) throws IOException {
        File file = new File(filePath);
        try (FileInputStream fis = new FileInputStream(file)) {
            byte[] wallpaperData = new byte[(int) file.length()];
            fis.read(wallpaperData);
            return wallpaperData;
        }
    }


    public static byte[] readByteArrayByUrl(String strUrl) {
        ByteArrayOutputStream baos = null;
        try {
            URL url = new URL(strUrl);
            BufferedImage image = ImageIO.read(url);

            // convert BufferedImage to byte array
            baos = new ByteArrayOutputStream();
            ImageIO.write(image, "jpg", baos);
            baos.flush();

            return baos.toByteArray();
        } catch (Exception ignored) {
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException ignored) {
                }
            }
        }
        return null;
    }

    public static String file2String(File file) {
        StringBuilder result = new StringBuilder();
        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            String s;
            while ((s = br.readLine()) != null) {
                result.append(System.lineSeparator()).append(s);
            }
            br.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result.toString();
    }

    public static void deleteFile(String fileName) {
        File file = new File(fileName);
        if (file.exists() && file.isFile()) {
            if (!file.delete()) {
                throw new IllegalStateException("delete file fail");
            }
        }
    }

    public static byte[] getImageFromNetByUrl(String strUrl) {
        try {
            URL url = new URL(strUrl);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5 * 1000);
            InputStream inStream = conn.getInputStream();// 通过输入流获取图片数据
            byte[] btImg = readInputStream(inStream);// 得到图片的二进制数据
            return btImg;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 从输入流中获取数据
     *
     * @param inStream 输入流
     * @return bytes
     * @throws Exception e
     */
    public static byte[] readInputStream(InputStream inStream) throws Exception {
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len = 0;
        while ((len = inStream.read(buffer)) != -1) {
            outStream.write(buffer, 0, len);
        }
        inStream.close();
        return outStream.toByteArray();
    }

    public static void delFolder(String folderPath) {
        try {
            delAllFile(folderPath); // 删除完里面所有内容
            String filePath = folderPath;
            filePath = filePath;
            File myFilePath = new File(filePath);
            myFilePath.delete(); // 删除空文件夹
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static boolean delAllFile(String path) {
        boolean flag = false;
        File file = new File(path);
        if (!file.exists()) {
            return flag;
        }
        if (!file.isDirectory()) {
            return flag;
        }
        String[] tempList = file.list();
        if (tempList == null) {
            return true;
        }
        File temp;
        for (String s : tempList) {
            if (path.endsWith(File.separator)) {
                temp = new File(path + s);
            } else {
                temp = new File(path + File.separator + s);
            }
            if (temp.isFile()) {
                temp.delete();
            }
            if (temp.isDirectory()) {
                delAllFile(path + "/" + s);// 先删除文件夹里面的文件
                delFolder(path + "/" + s);// 再删除空文件夹
                flag = true;
            }
        }
        return flag;
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

    public static String resolveUrl(String... pathMore) {
        StringBuilder path = new StringBuilder();
        for (String pstr : pathMore) {
            if (path.lastIndexOf("/") != (path.length() - 1)) {
                path.append("/");
            }
            path.append(pstr);
        }

        return path.toString();
    }


    public static void downloadFile(String urlStr, String path) throws IOException {
        DataInputStream dataInputStream = null;
        FileOutputStream fileOutputStream = null;
        ByteArrayOutputStream output = null;
        try {
            URL url = new URL(urlStr);
            dataInputStream = new DataInputStream(url.openStream());

            fileOutputStream = new FileOutputStream(new File(path));
            output = new ByteArrayOutputStream();

            byte[] buffer = new byte[1024];
            int length;

            while ((length = dataInputStream.read(buffer)) > 0) {
                output.write(buffer, 0, length);
            }
            fileOutputStream.write(output.toByteArray());
            dataInputStream.close();
            fileOutputStream.close();
        } finally {
            Closeables.closeAndSwallowIOExceptions(fileOutputStream, dataInputStream, output);
        }
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
