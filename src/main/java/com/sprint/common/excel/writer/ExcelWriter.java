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
import java.util.LinkedHashMap;
import java.util.Map;

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

    private final Map<ExcelCellStyle, CellStyle> styleMap = new LinkedHashMap<>();
    private final Map<String, ExcelSheetWriter> sheetMap = new LinkedHashMap<>();

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

    void putSheet(ExcelSheetWriter sheetExporter) {
        this.sheetMap.put(sheetExporter.currentSheet().getSheetName(), sheetExporter);
    }

    public ExcelSheetWriter createSheet() {
        String defaultSheetName = getDefaultSheetName();
        return createSheet(defaultSheetName, 0);
    }

    public ExcelSheetWriter createSheet(int maxRow) {
        String defaultSheetName = getDefaultSheetName();
        return createSheet(defaultSheetName, maxRow);
    }

    public ExcelSheetWriter createSheet(String sheetName) {
        return createSheet(sheetName, 0);
    }

    public ExcelSheetWriter createSheet(String sheetName, int maxRow) {
        ExcelSheetWriter excelSheetWriter =
                new ExcelSheetWriter(ExcelWriter.this, Excels.createSheet(ExcelWriter.this.getWorkbook(), sheetName), maxRow);
        this.putSheet(excelSheetWriter);
        return excelSheetWriter;
    }

    public ExcelSheetWriter getSheetWriterByName(String sheetName) {
        return this.sheetMap.get(sheetName);
    }

    public ExcelSheetWriter getCurrentDefaultSheetWriter() {
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

    private volatile ExcelDataExporter excelDataExporter;

    public ExcelDataExporter exporter() {
        if (excelDataExporter == null) {
            excelDataExporter = new ExcelDataExporter(this);
        }
        return excelDataExporter;
    }
}