package com.sprint.common.excel.util;

import com.sprint.common.excel.data.ExcelCell;
import com.sprint.common.excel.data.ExcelCellStyle;
import com.sprint.common.excel.data.ExcelType;
import com.sprint.common.excel.writer.Excelable;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.*;
import java.util.function.Supplier;


/**
 * excel工具
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2019年10月17日
 */
public class Excels {

    private static final Logger log = LoggerFactory.getLogger(Excels.class);

    private static final int DEF_ROW_ACCESS_WINDOW_SIZE = 10000;
    public static final String CLIENT_ABORT_EXCEPTION = "ClientAbortException";
    public static final int GBK_LEN = 24;

    public static boolean isExcelFile(InputStream inputStream) {
        return FileUtils.isExcel(inputStream);
    }

    public static ExcelType getExcelType(InputStream inputStream) {
        if (!isExcelFile(inputStream)) {
            return null;
        } else {
            try {
                FileType fileType = FileType.valueOf(inputStream);
                return Objects.equals(FileType.OLE2, fileType) ? ExcelType.XLS : ExcelType.XLSX;
            } catch (IOException ignored) {
                return null;
            }
        }
    }

    public static ExcelType getExcelTypeByFileName(String fileName) {
        if (fileName == null) {
            return null;
        } else {
            ExcelType[] var1 = ExcelType.values();
            for (ExcelType excelType : var1) {
                if (fileName.endsWith(excelType.getSuffix())) {
                    return excelType;
                }
            }
            return null;
        }
    }

    public static String getFileNameWithoutSuffix(String fileName) {
        ExcelType excelType = getExcelTypeByFileName(fileName);
        return excelType == null ? fileName : fileName.substring(0, fileName.length() - excelType.getSuffix().length());
    }

    public static boolean isExcelFileName(String fileName) {
        return getExcelTypeByFileName(fileName) != null;
    }

    public static String validateFileName(String fileName, Workbook wb) {
        return validateFileName(fileName, wb, null);
    }

    public static String validateFileName(String fileName, Workbook wb, Supplier<String> nameGen) {
        String suffix = getFileSuffix(wb);
        if (Miscs.isBlank(fileName)) {
            return nameGen == null ? System.currentTimeMillis() + suffix : validateFileName(nameGen.get(), wb);
        } else {
            return fileName.endsWith(suffix) ? fileName : fileName + suffix;
        }
    }

    public static String getFileSuffix(Workbook wb) {
        return ExcelType.fromWorkbook(wb).getSuffix();
    }

    public static String getContentType(Workbook wb) {
        return ExcelType.fromWorkbook(wb).getContentType();
    }

    public static Workbook createWorkbook() {
        return new HSSFWorkbook();
    }

    public static Workbook createLargeWorkbook() {
        return createLargeWorkbook(DEF_ROW_ACCESS_WINDOW_SIZE);
    }

    public static Workbook createLargeWorkbook(int rowAccessWindowSize) {
        return new SXSSFWorkbook(rowAccessWindowSize);
    }

    public static void exportToExcel(OutputStream output, Workbook wb) {
        try {
            wb.write(output);
        } catch (IOException ex) {
            String simplename = ex.getClass().getSimpleName();
            if (CLIENT_ABORT_EXCEPTION.equals(simplename)) {
                log.warn("Error while exporting data.And it's a ClientAbortException");
            }
        }
    }

    public static Sheet createSheet(Workbook wb, String sheetName) {
        Sheet sheet;
        if (sheetName != null) {
            sheet = wb.createSheet(sheetName);
        } else {
            sheet = wb.createSheet();
        }

        return sheet;
    }

    public static <A> int exportHeadRow2Sheet(Sheet sheet, Excelable<A> exporter) {
        int maxColumnNum = 0;
        ExcelCell[] cells = exporter.exportRowName();
        for (ExcelCell excelCell : cells) {
            maxColumnNum = Math.max(maxColumnNum, excelCell.getColumnNumOffset() + excelCell.getColumnSize());
        }

        for (int i = 0; i < maxColumnNum; ++i) {
            sheet.setColumnWidth(i, 5120);
        }

        Map<ExcelCellStyle, CellStyle> styleMap = new HashMap<>(16);
        return createRow(sheet, 0, exporter.exportRowName(), styleMap);
    }

    public static <A> int exportDaraRow2Sheet(Sheet sheet, int rowNum, List<A> datas, Excelable<A> exporter) {
        Map<ExcelCellStyle, CellStyle> styleMap = new HashMap<>(16);
        return exportDaraRow2Sheet(sheet, rowNum, datas, exporter, styleMap);
    }

    public static <A> int exportDaraRow2Sheet(Sheet sheet, int rowNum, List<A> datas, Excelable<A> exporter,
                                              Map<ExcelCellStyle, CellStyle> styleMap) {
        int rowNumAnchor = rowNum;

        for (A a : datas) {
            rowNumAnchor = createRow(sheet, rowNumAnchor, exporter.exportRowValue(a), styleMap);
        }

        return rowNumAnchor;
    }

    public static int createRow(Sheet sheet, int rowNumAnchor, ExcelCell[] values,
                                Map<ExcelCellStyle, CellStyle> styleMap) {
        int nextRowNum = rowNumAnchor;
        Workbook wb = sheet.getWorkbook();
        Drawing<?> p = sheet.createDrawingPatriarch();

        Map<Integer, List<ExcelCell>> mapByRowNum = new HashMap<>(16);

        for (ExcelCell value : values) {
            mapByRowNum.computeIfAbsent(value.getRowNumOffset(), ArrayList::new).add(value);
            nextRowNum = Math.max(rowNumAnchor + value.getRowNumOffset() + value.getRowSize(), nextRowNum);
        }

        List<Integer> rowNums = new ArrayList<>(mapByRowNum.keySet());
        Collections.sort(rowNums);
        rowNums.forEach((i) -> {
            int rowNum = rowNumAnchor + i;
            Row row = sheet.createRow(rowNum);

            for (ExcelCell excelCell : mapByRowNum.get(i)) {
                int columnNum = excelCell.getColumnNumOffset();
                Cell cell = row.createCell(columnNum);
                fitCellValueByCellType(cell, excelCell);

                if (Miscs.isNotBlank(excelCell.getComment())) {
                    cell.setCellComment(createCellComment(p, excelCell, (short) columnNum, rowNum));
                }

                ExcelCellStyle excelCellStyle = excelCell.getCellStyle();

                if (excelCellStyle != null) {
                    CellStyle cellStyle = styleMap.get(excelCellStyle);
                    if (cellStyle == null) {
                        cellStyle = wb.createCellStyle();
                        styleMap.put(excelCellStyle, cellStyle);
                        Font font = null;
                        if (excelCellStyle.getAlignment() != null) {
                            cellStyle.setAlignment(HorizontalAlignment.forInt(excelCellStyle.getAlignment()));
                        }

                        if (excelCellStyle.getVerticalAlignment() != null) {
                            cellStyle
                                    .setVerticalAlignment(VerticalAlignment.forInt(excelCellStyle.getVerticalAlignment()));
                        }

                        if (excelCellStyle.getFontColor() != null) {
                            font = wb.createFont();

                            font.setColor(excelCellStyle.getFontColor());
                        }

                        if (excelCellStyle.getBold() != null && excelCellStyle.getBold()) {
                            if (font == null) {
                                font = wb.createFont();
                            }
                            font.setBold(excelCellStyle.getBold());
                        }

                        if (font != null) {
                            cellStyle.setFont(font);
                        }
                    }

                    cell.setCellStyle(cellStyle);
                }

                if (excelCell.getRowSize() != 1 || excelCell.getColumnSize() != 1) {
                    sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum + excelCell.getRowSize() - 1, columnNum,
                            columnNum + excelCell.getColumnSize() - 1));
                }
            }
        });
        return nextRowNum;
    }

    private static void fitCellValueByCellType(Cell cell, ExcelCell value) {
        Integer cellType = value.getCellType();
        Object val = value.getValue();
        if (cellType == null) {
            cellType = 1;
        }

        if (cellType == 0) {
            if (val instanceof Number) {
                cell.setCellValue(Double.parseDouble(val.toString()));
            } else {
                cell.setCellValue(0.0D);
            }
        } else {
            cell.setCellValue(val.toString());
        }

    }

    public static Comment createCellComment(Drawing<?> draw, ExcelCell value, short col, int row) {
        short col2;
        int row2 = row + 1;
        int len = 0;

        try {
            len = value.getComment().getBytes("GBK").length;
        } catch (UnsupportedEncodingException var8) {
            var8.printStackTrace();
        }

        if (len > GBK_LEN) {
            col2 = (short) (col + 3);
            row2 = row + (len - 1) / 24 + 1;
        } else {
            col2 = (short) (col + (len + 2) / 9 + 1);
        }

        Comment comment = draw.createCellComment(new HSSFClientAnchor(0, 0, 0, 127, col, row, col2, row2));
        comment.setString(new HSSFRichTextString(value.getComment()));
        comment.setVisible(value.getIsCommentVisible());
        return comment;
    }

    public static void setSimpleCellColumnOffset(ExcelCell[] excelCells) {
        for (int i = 0; i < excelCells.length; ++i) {
            excelCells[i].setColumnNumOffset(i);
        }
    }

}