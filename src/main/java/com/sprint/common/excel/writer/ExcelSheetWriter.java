package com.sprint.common.excel.writer;

import com.sprint.common.excel.data.ExcelCellStyle;
import com.sprint.common.excel.data.ExcelType;
import com.sprint.common.excel.util.Excels;
import com.sprint.common.excel.util.Safes;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * 写Excel sheet
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2019年10月17日
 */
public class ExcelSheetWriter {

    public final List<Sheet> sheets = new ArrayList<>();

    private final ExcelWriter excelWriter;
    private final Map<ExcelCellStyle, CellStyle> styleMap;

    private int currentRowNum = 0;
    private String sheetName;
    private int maxRow;
    private int index = 0;

    ExcelSheetWriter(ExcelWriter excelWriter, Sheet sheet) {
        this(excelWriter, sheet, 0);
    }

    ExcelSheetWriter(ExcelWriter excelWriter, Sheet sheet, int maxRow) {
        this.excelWriter = excelWriter;
        this.sheets.add(sheet);
        this.styleMap = excelWriter.getStyleMap();
        this.maxRow = maxRow;
        init();
    }

    private void init() {
        Sheet currentSheet = currentSheet();
        this.sheetName = currentSheet.getSheetName();
        ExcelType excelType = ExcelType.fromWorkbook(currentSheet.getWorkbook());
        if (this.maxRow == 0) {
            this.maxRow = excelType.getMaxRow();
        }
    }

    public ExcelDataExporter exporter() {
        return new ExcelDataExporter(excelWriter);
    }

    public String getSheetName() {
        return this.sheetName;
    }

    void addSheet(Sheet sheet) {
        this.sheets.add(sheet);
    }

    Sheet currentSheet() {
        return Safes.last(this.sheets);
    }

    public int getCurrentRowNum() {
        return currentRowNum;
    }

    void setCurrentRowNum(int currentRowNum) {
        this.currentRowNum = currentRowNum;
    }

    public int getMaxRow() {
        return maxRow;
    }

    public <T> RowWriter<T> excelable(Excelable<T> exporter) {
        return new RowWriter<>(exporter);
    }

    public <T> RowWriter<T> empty() {
        return excelable(Excelables.none());
    }

    public <T> RowWriter<T> of(Map<String, String> mapper, Class<T> rowType) {
        return excelable(Excelables.of(mapper));
    }


    public <T> RowWriter<T> ofNumber(Map<String, String> mapper) {
        return excelable(Excelables.ofNumber(mapper));
    }

    public <T> RowWriter<T> of(String[] headers, String[] propertys) {
        return excelable(Excelables.of(headers, propertys));
    }

    public <T> RowWriter<T> ofNumber(String[] headers, String[] propertys) {
        return excelable(Excelables.ofNumber(headers, propertys));
    }

    public <T> RowWriter<T> of(String[] headers) {
        return excelable(Excelables.of(headers, headers));
    }

    public <T> RowWriter<T> of(String[] headers, String[] propertys,
                               BiFunction<T, String, String> function) {
        return excelable(Excelables.of(headers, propertys, function));
    }

    public <T> RowWriter<T> of(Class<T> xCelBeanClass) {
        return excelable(Excelables.of(xCelBeanClass));
    }

    public class RowWriter<T> {

        public final Excelable<T> exporter;

        public RowWriter(Excelable<T> exporter) {
            this.exporter = exporter;
            begin();
        }

        private void begin() {
            setCurrentRowNum(Excels.exportHeadRow2Sheet(currentSheet(), this.exporter));
        }

        private void nextSheet() {
            ++ExcelSheetWriter.this.index;
            addSheet(currentSheet().getWorkbook().createSheet(getSheetName() + ExcelSheetWriter.this.index));
            setCurrentRowNum(0);
            excelWriter.putSheet(ExcelSheetWriter.this);
            this.begin();
        }

        public ExcelDataExporter addRowList(List<T> dataList) {
            if (dataList != null) {
                if (getCurrentRowNum() + dataList.size() >= getMaxRow()) {
                    int splitIndex = getMaxRow() - getCurrentRowNum();
                    setCurrentRowNum(Excels.exportDaraRow2Sheet(currentSheet(), getCurrentRowNum(), dataList.subList(0, splitIndex),
                            this.exporter, ExcelSheetWriter.this.styleMap));
                    nextSheet();
                    this.addRowList(dataList.subList(splitIndex, dataList.size() - 1));
                } else {
                    setCurrentRowNum(Excels.exportDaraRow2Sheet(currentSheet(), getCurrentRowNum(), dataList, this.exporter));
                }
            }
            return excelWriter.exporter();
        }

        public ExcelDataExporter addRows(T... t) {
            addRowList(Arrays.asList(t));
            return excelWriter.exporter();
        }
    }
}