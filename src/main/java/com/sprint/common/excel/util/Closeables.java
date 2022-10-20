package com.sprint.common.excel.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;
import java.io.IOException;


/**
 * 可关闭的
 */
public final class Closeables {

    private static final Logger logger = LoggerFactory.getLogger(Closeables.class);

    private Closeables() {
    }

    public static void close(Closeable... closeables) throws IOException {
        for (Closeable c : closeables) {
            if (c != null) {
                c.close();
            }
        }
    }

    public static void closeAndSwallowIOExceptions(Closeable... closeables) {
        for (Closeable c : closeables) {
            if (c != null) {
                try {
                    c.close();
                } catch (IOException ex) {
                    logger.error("Encountered exception closing closeable", ex);
                }
            }
        }
    }
}
