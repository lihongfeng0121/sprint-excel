package com.sprint.common.excel.util;

import java.io.Closeable;
import java.io.IOException;


/**
 * 可关闭的
 */
public final class Closeables {

    private Closeables() {
    }

    public static void close(Closeable... closeables) throws IOException {
        for (Closeable c : closeables) {
            if (c != null) {
                c.close();
            }
        }
    }
}
