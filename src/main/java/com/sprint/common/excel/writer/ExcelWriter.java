package com.sprint.common.excel.writer;

import com.sprint.common.excel.data.ExcelCellStyle;
import com.sprint.common.excel.util.Excels;
import com.sprint.common.excel.util.Miscs;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * excel 写入
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2019年09月09日
 */
public class ExcelWriter {

    private final String fileName;
    private final Workbook workbook;

    private final Map<ExcelCellStyle, CellStyle> styleMap = new HashMap<>();
    private final Map<String, ExcelSheetWriter<?>> sheetMap = new HashMap<>();

    private static final String DEFAULT_SHEET_NAME = "sheet";

    private ExcelWriter(String fileName, Workbook workbook) {
        this.fileName = Excels.validateFileName(fileName, workbook);
        this.workbook = workbook;
    }

    public static ExcelWriter xls(String fileName) {
        return new ExcelWriter(fileName, Excels.createWorkbook());
    }

    public static ExcelWriter xlsx(String fileName) {
        return new ExcelWriter(fileName, Excels.createLargeWorkbook());
    }

    public static ExcelWriter xlsx(String fileName, int rowAccessWindowSize) {
        return new ExcelWriter(fileName, Excels.createLargeWorkbook(rowAccessWindowSize));
    }

    public void export(OutputStream output) throws IOException {
        if (Miscs.isEmpty(sheetMap)) {
            workbook.createSheet();
        }
        Excels.exportToExcel(output, this.workbook);
        close();
    }

    private void close() throws IOException {
        try {
            Method method = Workbook.class.getMethod("close");
            if (!method.isAccessible()) {
                method.setAccessible(true);
            }
            method.invoke(this.workbook);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException ignored) {
        }
    }

    void putSheet(ExcelSheetWriter<?> sheetExporter) {
        this.sheetMap.put(sheetExporter.getSheet().getSheetName(), sheetExporter);
    }

    public SheetCreator createSheet() {
        return new SheetCreator(getDefaultSheetName(), 0);
    }

    public SheetCreator createSheet(int maxRow) {
        return new SheetCreator(getDefaultSheetName(), maxRow);
    }

    public SheetCreator createSheet(String sheetName) {
        return new SheetCreator(sheetName, 0);
    }

    public SheetCreator createSheet(String sheetName, int maxRow) {
        return new SheetCreator(sheetName, maxRow);
    }

    public ExcelSheetWriter<?> getSheetWriterByName(String sheetName) {
        return this.sheetMap.get(sheetName);
    }

    public ExcelSheetWriter<?> getCurrentDefaultSheetWriter() {
        return this.sheetMap.get(getDefaultSheetName());
    }

    private String getDefaultSheetName() {
        return DEFAULT_SHEET_NAME + sheetMap.size();
    }

    public String getFileName() {
        return this.fileName;
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    Map<ExcelCellStyle, CellStyle> getStyleMap() {
        return this.styleMap;
    }


    public class SheetCreator {

        private final String sheetName;
        private final int maxRow;

        public SheetCreator(String sheetName, int maxRow) {
            this.sheetName = sheetName;
            this.maxRow = maxRow;
        }


        public <T> ExcelSheetWriter<T> excelable(Excelable<T> exporter) {
            ExcelSheetWriter<T> sheetExporter =
                    new ExcelSheetWriter<>(ExcelWriter.this, Excels.createSheet(ExcelWriter.this.getWorkbook(), sheetName), exporter, maxRow);
            ExcelWriter.this.putSheet(sheetExporter);
            return sheetExporter;
        }

        public <T> ExcelSheetWriter<T> empty() {
            ExcelSheetWriter<T> sheetExporter =
                    new ExcelSheetWriter<>(ExcelWriter.this, Excels.createSheet(ExcelWriter.this.getWorkbook(), sheetName), Excelables.none(), maxRow);
            ExcelWriter.this.putSheet(sheetExporter);
            return sheetExporter;
        }

        public <T> ExcelSheetWriter<T> of(Map<String, String> mapper) {
            return excelable(Excelables.of(mapper));
        }


        public <T> ExcelSheetWriter<T> ofNumber(Map<String, String> mapper) {
            return excelable(Excelables.ofNumber(mapper));
        }

        public <T> ExcelSheetWriter<T> of(String[] headers, String[] propertys) {
            return excelable(Excelables.of(headers, propertys));
        }

        public <T> ExcelSheetWriter<T> ofNumber(String[] headers, String[] propertys) {
            return excelable(Excelables.ofNumber(headers, propertys));
        }

        public <T> ExcelSheetWriter<T> of(String[] headers) {
            return excelable(Excelables.of(headers, headers));
        }

        public <T> ExcelSheetWriter<T> of(String[] headers, String[] propertys,
                                          BiFunction<T, String, String> function) {
            return excelable(Excelables.of(headers, propertys, function));
        }

        public <T> ExcelSheetWriter<T> of(Class<T> xCelBeanClass) {
            return excelable(Excelables.of(xCelBeanClass));
        }

    }

}