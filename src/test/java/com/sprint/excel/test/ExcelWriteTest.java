package com.sprint.excel.test;

import com.sprint.common.excel.writer.ExcelSheetWriter;
import com.sprint.common.excel.writer.ExcelWriter;
import org.junit.Test;

import java.io.IOException;
import java.util.Collections;
import java.util.Map;

/**
 * @author hongfeng.li
 * @since 2022/10/21
 */
public class ExcelWriteTest {

    @Test
    public void test() throws IOException {
        ExcelWriter test21 = ExcelWriter.xlsx("test5");
        ExcelSheetWriter test2 = test21.createSheet();
        ExcelSheetWriter test3 = test21.createSheet();
        test21.exporter().export("/Users/lihongfeng/Desktop/");
    }

}
