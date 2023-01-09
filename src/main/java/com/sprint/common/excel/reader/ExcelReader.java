package com.sprint.common.excel.reader;

import com.sprint.common.excel.data.Sheet;

import java.util.List;

/**
 * excel reader
 *
 * @author hongfeng-li
 * @since 2018/11/2910:31
 */
public interface ExcelReader {

    /**
     * 处理
     *
     * @return sheets
     * @throws Exception e
     */
    List<Sheet> process() throws Exception;

}
