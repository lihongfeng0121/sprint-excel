package com.sprint.common.excel.reader;

import com.sprint.common.converter.BeanConverter;
import com.sprint.common.converter.conversion.nested.bean.introspection.CachedIntrospectionResults;
import com.sprint.common.converter.conversion.nested.bean.introspection.PropertyAccess;
import com.sprint.common.excel.data.*;
import com.sprint.common.excel.reader.excel03.Excel03Reader;
import com.sprint.common.excel.reader.excel07.Excel07Reader;
import com.sprint.common.excel.util.Closeables;
import com.sprint.common.excel.util.Excels;
import com.sprint.common.excel.util.Miscs;
import com.sprint.common.excel.util.Safes;
import org.apache.poi.ooxml.util.PackageHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Excel 读工具
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2019年10月17日
 */
public class UtilExcelReader {

    private static final Logger logger = LoggerFactory.getLogger(UtilExcelReader.class);

    public static final int DEFAULT_TITLE_ROW_NUMBER = 0;

    public static final int DEFAULT_DATA_ROW_NUMBER = 1;

    /**
     * 读Excel
     *
     * @param fin 输入流
     * @return 行
     * @throws Exception e
     */
    public static List<XRow> readExcel(InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel(inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param fin       输入流
     * @param excelType excel类型
     * @return 行
     * @throws Exception e
     */
    public static List<XRow> readExcel(InputStream fin, ExcelType excelType) throws Exception {
        try {
            Map<String, List<XRow>> sheetMap = readExcelSheets(fin, excelType);

            List<XRow> allXRows = new ArrayList<>();

            for (List<XRow> xRows : sheetMap.values()) {
                if (Miscs.isNotEmpty(xRows)) {
                    allXRows.addAll(xRows);
                }
            }

            return allXRows;
        } finally {
            Closeables.close(fin);
        }
    }

    /**
     * 读Excel
     *
     * @param sheetName sheetName
     * @param fin       输入流
     * @return 行
     * @throws Exception e
     */
    public static List<XRow> readExcelSheet(String sheetName, InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcelSheet(sheetName, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param sheetName sheetName
     * @param fin       输入流
     * @param excelType excel类型
     * @return 行
     * @throws Exception e
     */
    public static List<XRow> readExcelSheet(String sheetName, InputStream fin, ExcelType excelType) throws Exception {
        try {
            return readExcelSheets(fin, excelType).get(sheetName);
        } finally {
            Closeables.close(fin);
        }
    }

    /**
     * 读Excel
     *
     * @param sheetNumber sheetNumber
     * @param fin         输入流
     * @return 行
     * @throws Exception e
     */
    public static List<XRow> readExcelSheet(int sheetNumber, InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcelSheet(sheetNumber, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param sheetNumber sheetNumber
     * @param fin         输入流
     * @param excelType   excel类型
     * @return 行
     * @throws Exception e
     */
    public static List<XRow> readExcelSheet(int sheetNumber, InputStream fin, ExcelType excelType) throws Exception {
        try {
            Map<String, List<XRow>> allSheetRows = readExcelSheets(fin, excelType);
            if (Miscs.isEmpty(allSheetRows)) {
                return Collections.emptyList();
            }
            String sheetName = Safes.at(allSheetRows.keySet(), sheetNumber);
            return allSheetRows.get(sheetName);
        } finally {
            Closeables.close(fin);
        }
    }

    /**
     * 读Excel
     *
     * @param fin 输入流
     * @return 行
     * @throws Exception e
     */
    public static Map<String, List<XRow>> readExcelSheets(InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcelSheets(inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param fin       输入流
     * @param excelType excel类型
     * @return 行
     * @throws Exception e
     */
    public static Map<String, List<XRow>> readExcelSheets(InputStream fin, ExcelType excelType) throws Exception {
        ExcelReader reader;
        try {
            switch (excelType) {
                case XLS:
                    reader = new Excel03Reader(fin);
                    break;
                case XLSX:
                    reader = new Excel07Reader(PackageHelper.open(fin));
                    break;
                default:
                    throw new IllegalArgumentException("excel extension illegal.");
            }

            Map<String, List<XRow>> sheetMap = reader.process().stream()
                    .collect(Collectors.toMap(Sheet::getSheetName, Sheet::getRows, (o1, o2) -> o1, LinkedHashMap::new));

            return Miscs.rewriteVal(sheetMap,
                    (val) -> val.stream().filter(row -> !row.isEmpty()).collect(Collectors.toList()));
        } finally {
            Closeables.close(fin);
        }
    }

    /**
     * 读Excel
     *
     * @param titleRowNumber  标题行号
     * @param dateStartNumber 数据行号
     * @param fin             输入流
     * @return 行
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(int titleRowNumber, int dateStartNumber, InputStream fin)
            throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2XTDRow(titleRowNumber, dateStartNumber, inputStreams[0],
                Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param titleRowNumber  标题行号
     * @param dateStartNumber 数据行号
     * @param fin             输入流
     * @param excelType       excel类型
     * @return 行
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(int titleRowNumber, int dateStartNumber, InputStream fin,
                                                ExcelType excelType) throws Exception {
        List<XTDRow> list = new LinkedList<>();
        Map<String, List<XRow>> allSheetRows = readExcelSheets(fin, excelType);
        if (Miscs.isEmpty(allSheetRows)) {
            return Collections.emptyList();
        }

        for (List<XRow> rows : allSheetRows.values()) {
            // 获取到第一行，也即表格列名称
            list.addAll(xRows2XTDRow(rows, titleRowNumber, dateStartNumber));
        }

        return list;
    }

    /**
     * 读Excel
     *
     * @param sheetName       sheetName
     * @param titleRowNumber  标题行号
     * @param dateStartNumber 数据行号
     * @param fin             输入流
     * @return List
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(String sheetName, int titleRowNumber, int dateStartNumber,
                                                InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2XTDRow(sheetName, titleRowNumber, dateStartNumber, inputStreams[0],
                Excels.getExcelType(inputStreams[1]));
    }


    public static List<XTDRow> readExcel2XTDRow(String sheetName, int titleRowNumber, int dateStartNumber,
                                                InputStream fin, ExcelType excelType) throws Exception {

        if (Miscs.isEmpty(sheetName)) {
            return Collections.emptyList();
        }

        Map<String, List<XRow>> allSheetRows = readExcelSheets(fin, excelType);
        if (Miscs.isEmpty(allSheetRows)) {
            return Collections.emptyList();
        }

        List<XRow> rows = allSheetRows.get(sheetName);

        if (rows == null) {
            logger.warn("UtilExcelReader#readExcel2Map sheetName[] not exit!", sheetName);
            return Collections.emptyList();
        }

        return xRows2XTDRow(rows, titleRowNumber, dateStartNumber);
    }

    /**
     * 读Excel
     *
     * @param sheetName sheetName
     * @param fin       输入流
     * @return List
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(String sheetName, InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2XTDRow(sheetName, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param sheetName sheetName
     * @param fin       输入流
     * @param excelType excel类型
     * @return List
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(String sheetName, InputStream fin, ExcelType excelType)
            throws Exception {
        return readExcel2XTDRow(sheetName, DEFAULT_TITLE_ROW_NUMBER, DEFAULT_DATA_ROW_NUMBER, fin, excelType);
    }

    /**
     * 读Excel
     *
     * @param fin 输入流
     * @return List
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2XTDRow(inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param fin       输入流
     * @param excelType excel类型
     * @return List
     * @throws Exception e
     */
    public static List<XTDRow> readExcel2XTDRow(InputStream fin, ExcelType excelType) throws Exception {
        return readExcel2XTDRow(DEFAULT_TITLE_ROW_NUMBER, DEFAULT_DATA_ROW_NUMBER, fin, excelType);
    }

    public static List<Map<String, XCol>> readExcel2Map(String sheetName, InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Map(sheetName, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    public static List<Map<String, XCol>> readExcel2Map(String sheetName, InputStream fin, ExcelType excelType)
            throws Exception {
        return readExcel2Map(sheetName, DEFAULT_TITLE_ROW_NUMBER, DEFAULT_DATA_ROW_NUMBER, fin, excelType);
    }

    public static List<Map<String, XCol>> readExcel2Map(String sheetName, int titleRowNumber, int dateStartNumber,
                                                        InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Map(sheetName, titleRowNumber, dateStartNumber, inputStreams[0],
                Excels.getExcelType(inputStreams[1]));
    }

    public static List<Map<String, XCol>> readExcel2Map(String sheetName, int titleRowNumber, int dateStartNumber,
                                                        InputStream fin, ExcelType excelType) throws Exception {

        if (Miscs.isEmpty(sheetName)) {
            return Collections.emptyList();
        }

        Map<String, List<XRow>> allSheetRows = readExcelSheets(fin, excelType);
        if (Miscs.isEmpty(allSheetRows)) {
            return Collections.emptyList();
        }

        List<XRow> rows = allSheetRows.get(sheetName);

        if (rows == null) {
            logger.warn("UtilExcelReader#readExcel2Map sheetName[] not exit!", sheetName);
            return Collections.emptyList();
        }

        return xRows2Map(rows, titleRowNumber, dateStartNumber);
    }

    public static List<Map<String, XCol>> readExcel2Map(int sheetNumber, InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Map(sheetNumber, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    public static List<Map<String, XCol>> readExcel2Map(int sheetNumber, InputStream fin, ExcelType excelType)
            throws Exception {
        return readExcel2Map(sheetNumber, DEFAULT_TITLE_ROW_NUMBER, DEFAULT_DATA_ROW_NUMBER, fin, excelType);
    }

    public static List<Map<String, XCol>> readExcel2Map(int sheetNumber, int titleRowNumber, int dateStartNumber,
                                                        InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Map(sheetNumber, titleRowNumber, dateStartNumber, inputStreams[0],
                Excels.getExcelType(inputStreams[1]));
    }

    public static List<Map<String, XCol>> readExcel2Map(int sheetNumber, int titleRowNumber, int dateStartNumber,
                                                        InputStream fin, ExcelType excelType) throws Exception {
        Map<String, List<XRow>> allSheetRows = readExcelSheets(fin, excelType);
        if (Miscs.isEmpty(allSheetRows)) {
            return Collections.emptyList();
        }

        String sheetName = Safes.at(allSheetRows.keySet(), sheetNumber);

        if (Miscs.isEmpty(sheetName)) {
            logger.warn("UtilExcelReader#readExcel2Map sheet[] not exit!", sheetNumber);
            return Collections.emptyList();
        }

        return xRows2Map(allSheetRows.get(sheetName), titleRowNumber, dateStartNumber);
    }

    /**
     * 读Excel
     *
     * @param dateStartNumber dateStartNumber
     * @param titleRowNumber  titleRowNumber
     * @param fin             输入流
     * @return 行List
     * @throws Exception e
     */
    public static List<Map<String, XCol>> readExcel2Map(int titleRowNumber, int dateStartNumber, InputStream fin)
            throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Map(titleRowNumber, dateStartNumber, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param dateStartNumber dateStartNumber
     * @param titleRowNumber  titleRowNumber
     * @param fin             输入流
     * @param excelType       excel类型
     * @return 行List
     * @throws Exception e
     */
    public static List<Map<String, XCol>> readExcel2Map(int titleRowNumber, int dateStartNumber, InputStream fin,
                                                        ExcelType excelType) throws Exception {
        List<Map<String, XCol>> list = new LinkedList<>();
        Map<String, List<XRow>> allSheetRows = readExcelSheets(fin, excelType);
        if (Miscs.isEmpty(allSheetRows)) {
            return Collections.emptyList();
        }

        for (List<XRow> rows : allSheetRows.values()) {
            // 获取到第一行，也即表格列名称
            list.addAll(xRows2Map(rows, titleRowNumber, dateStartNumber));
        }

        return list;
    }

    /**
     * 读Excel
     *
     * @param fin 输入流
     * @return 行List
     * @throws Exception e
     */
    public static List<Map<String, XCol>> readExcel2Map(InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Map(inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    /**
     * 读Excel
     *
     * @param fin       输入流
     * @param excelType excel类型
     * @return 行List
     * @throws Exception e
     */
    public static List<Map<String, XCol>> readExcel2Map(InputStream fin, ExcelType excelType) throws Exception {
        return readExcel2Map(DEFAULT_TITLE_ROW_NUMBER, DEFAULT_DATA_ROW_NUMBER, fin, excelType);
    }

    public static <T> List<T> readExcel2Bean(Class<T> beanType, InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Bean(beanType, inputStreams[0], Excels.getExcelType(inputStreams[1]));
    }

    public static <T> List<T> readExcel2Bean(Class<T> beanType, InputStream fin, ExcelType excelType) throws Exception {
        return readExcel2Bean(DEFAULT_TITLE_ROW_NUMBER, DEFAULT_DATA_ROW_NUMBER, beanType, fin, excelType);
    }

    public static <T> List<T> readExcel2Bean(int titleRowNumber, int dateStartNumber, Class<T> beanType,
                                             InputStream fin) throws Exception {
        InputStream[] inputStreams = Miscs.copyInputStream(fin, 2);
        return readExcel2Bean(titleRowNumber, dateStartNumber, beanType, inputStreams[0],
                Excels.getExcelType(inputStreams[1]));
    }

    public static <T> List<T> readExcel2Bean(int titleRowNumber, int dateStartNumber, Class<T> beanType,
                                             InputStream fin, ExcelType excelType) throws Exception {
        List<Map<String, XCol>> mapList = readExcel2Map(titleRowNumber, dateStartNumber, fin, excelType);
        if (Miscs.isEmpty(mapList)) {
            return Collections.emptyList();
        }
        CachedIntrospectionResults introspectionResults = CachedIntrospectionResults.forClass(beanType);
        PropertyAccess[] writePropertyAccess = introspectionResults.getWritePropertyAccess();
        Map<String, String> cell2FieldMap = Stream.of(writePropertyAccess).collect(Collectors.toMap(item -> {
            XCell xCell = item.getAnnotation(XCell.class);
            return Optional.ofNullable(xCell)
                    .map(XCell::name).filter(Miscs::isNotBlank).orElse(item.getName());
        }, PropertyAccess::getName, (o1, o2) -> o2));
        return mapList.stream().map(map -> Miscs.rewrite(map, cell2FieldMap::get, XCol::getValue))
                .map(item -> BeanConverter.convert(item, beanType)).collect(Collectors.toList());
    }

    public static List<Map<String, XCol>> xRows2Map(List<XRow> xRows, int titleRowNumber, int dateStartNumber) {
        return xRows2XTDRow(xRows, titleRowNumber, dateStartNumber).stream().map(XTDRow::getTitleColMap)
                .collect(Collectors.toList());
    }

    public static List<XTDRow> xRows2XTDRow(List<XRow> xRows, int titleRowNumber, int dateStartNumber) {
        if (Miscs.size(xRows) < dateStartNumber) {
            return Collections.emptyList();
        }

        List<XTDRow> result = new ArrayList<>(xRows.size() - dateStartNumber);
        // 获取到第一行，也即表格列名称
        XRow titles = xRows.get(titleRowNumber);
        // 遍历行数
        for (int i = dateStartNumber; i < xRows.size(); i++) {

            if (titleRowNumber == i) {
                continue;
            }

            XRow row = xRows.get(i);

            result.add(new XTDRow(titles, row));
        }

        return result;
    }
}