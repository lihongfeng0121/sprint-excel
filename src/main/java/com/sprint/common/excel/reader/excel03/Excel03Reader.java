package com.sprint.common.excel.reader.excel03;

import com.sprint.common.excel.data.Sheet;
import com.sprint.common.excel.data.XCol;
import com.sprint.common.excel.data.XRow;
import com.sprint.common.excel.reader.ExcelReader;
import com.sprint.common.excel.util.Closeables;
import com.sprint.common.excel.util.Safes;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Excel03Reader implements ExcelReader, HSSFListener {

    private SSTRecord sstRecord;

    private final Map<String, List<XRow>> rows = new LinkedHashMap<>();

    private FormatTrackingHSSFListener formatListener;

    private int totalRows;

    private InputStream fin;

    private int sheetNumber = 0;

    private List<XRow> currentSheet = new ArrayList<>();

    public Excel03Reader(InputStream fin) {
        this.fin = fin;
    }

    @Override
    public List<Sheet> process() throws IOException {
        InputStream din = null;
        try {
            // 空值单元监听器
            MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
            // 格式监听器
            formatListener = new FormatTrackingHSSFListener(listener);
            // create a new org.apache.poi.poifs.filesystem.Filesystem
            POIFSFileSystem poifs = new POIFSFileSystem(fin);
            // get the Workbook (excel part) stream in a InputStream
            din = poifs.createDocumentInputStream("Workbook");
            // construct out HSSFRequest object
            HSSFRequest req = new HSSFRequest();
            // lazy listen for ALL records with the listener shown above
            req.addListenerForAllRecords(formatListener);
            // create our event factory
            HSSFEventFactory factory = new HSSFEventFactory();
            // process our events based on the document input stream
            factory.processEvents(req, din);
            // once all the events are processed close our file input stream
        } finally {
            // and our document input stream (don't want to leak these!)
            Closeables.close(fin, din);
        }
        return this.getSheets();
    }

    private List<Sheet> getSheets() {
        int sheetNumber = 1;
        List<Sheet> sheets = new ArrayList<>(rows.size());
        for (Map.Entry<String, List<XRow>> entry : rows.entrySet()) {
            sheets.add(new Sheet(sheetNumber++, entry.getKey(), entry.getValue()));
        }
        return sheets;
    }

    /**
     * This method listens for incoming records and handles them as required.
     *
     * @param record The record that was found while reading.
     */
    @Override
    public void processRecord(Record record) {
        switch (record.getSid()) {
            // the BOFRecord can represent either the beginning of a sheet or the
            // workbook
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                    currentSheet = rows.get(Safes.at(rows.keySet(), sheetNumber++));
                    totalRows = currentSheet.size();
                }
                break;
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                String sheetName = bsr.getSheetname();
                rows.computeIfAbsent(sheetName, (key) -> new ArrayList<>());
                break;

            case RowRecord.sid:
                XRow xrow = new XRow(currentSheet.size() + 1);
                currentSheet.add(xrow);
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;
                String value = formatListener.formatNumberDateCell(numrec).trim();
                XRow row = currentSheet.get(getRowNum(numrec.getRow()));
                row.getCol().add(new XCol(value));
                break;
            // SSTRecords store a array of unique strings used in Excel.
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                XCol strCol = new XCol(sstRecord.getString(lrec.getSSTIndex()).getString());
                currentSheet.get(getRowNum(lrec.getRow())).getCol().add(strCol);
                break;
            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;
                currentSheet.get(getRowNum(brec.getRow())).getCol().add(XCol.EMPTY);
                break;
            case FormulaRecord.sid:
                FormulaRecord formulaRecord = (FormulaRecord) record;
                XCol formulaCol = new XCol(formatListener.formatNumberDateCell(formulaRecord));
                currentSheet.get(getRowNum(formulaRecord.getRow())).getCol().add(formulaCol);
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;
                XCol boolCol = new XCol(String.valueOf(berec.getBooleanValue()));
                currentSheet.get(getRowNum(berec.getRow())).getCol().add(boolCol);
                break;
            default:
                break;
        }
        // 处理空值单元格
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            currentSheet.get(getRowNum(mc.getRow())).getCol().add(XCol.EMPTY);
        }
    }

    private int getRowNum(int curRow) {
        return curRow + totalRows;
    }

    public Map<String, List<XRow>> getRows() {
        return rows;
    }

    public List<XRow> getRowsOfSheet(String sheetName) {
        return rows.get(sheetName);
    }
}
