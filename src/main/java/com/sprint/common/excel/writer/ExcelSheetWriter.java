package com.sprint.common.excel.writer;

import com.sprint.common.excel.data.ExcelCellStyle;
import com.sprint.common.excel.data.ExcelType;
import com.sprint.common.excel.util.Excels;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * 写Excel sheet
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2019年10月17日
 */
public class ExcelSheetWriter<T> {

    private ExcelWriter excelWriter;
    private Sheet sheet;
    private Excelable<T> exporter;
    private Map<ExcelCellStyle, CellStyle> styleMap;
    private int rowNum = 0;
    private String sheetName;
    private int maxRow;
    private int index = 0;

    ExcelSheetWriter(ExcelWriter excelWriter, Sheet sheet, Excelable<T> exporter) {
        this(excelWriter, sheet, exporter, 0);
    }

    ExcelSheetWriter(ExcelWriter excelWriter, Sheet sheet, Excelable<T> exporter, int maxRow) {
        this.excelWriter = excelWriter;
        this.sheet = sheet;
        this.exporter = exporter;
        this.styleMap = excelWriter.getStyleMap();
        this.maxRow = maxRow;
        this.init();
    }

    private void init() {
        this.sheetName = this.sheet.getSheetName();
        ExcelType excelType = ExcelType.fromWorkbook(this.sheet.getWorkbook());
        if (this.maxRow == 0) {
            this.maxRow = excelType.getMaxRow();
        }
        this.begin();
    }

    private void begin() {
        this.rowNum = Excels.exportHeadRow2Sheet(this.sheet, this.exporter);
    }

    public ExcelDataExporter addRow(List<T> dataList) {
        if (dataList != null) {
            if (this.rowNum + dataList.size() >= this.maxRow) {
                int splitIndex = this.maxRow - this.rowNum;
                this.rowNum = Excels.exportDaraRow2Sheet(this.sheet, this.rowNum, dataList.subList(0, splitIndex),
                        this.exporter, this.styleMap);
                this.nextSheet();
                this.addRow(dataList.subList(splitIndex, dataList.size() - 1));
            } else {
                this.rowNum = Excels.exportDaraRow2Sheet(this.sheet, this.rowNum, dataList, this.exporter);
            }
        }
        return new ExcelDataExporter(excelWriter);
    }

    public ExcelDataExporter addRow(T... t) {
        addRow(Arrays.asList(t));
        return new ExcelDataExporter(excelWriter);
    }

    private void nextSheet() {
        ++this.index;
        this.sheet = this.sheet.getWorkbook().createSheet(this.sheetName + this.index);
        this.rowNum = 0;
        this.excelWriter.putSheet(this);
        this.begin();
    }

    Sheet getSheet() {
        return this.sheet;
    }
}