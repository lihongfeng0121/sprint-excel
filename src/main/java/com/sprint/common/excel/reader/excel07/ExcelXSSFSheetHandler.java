package com.sprint.common.excel.reader.excel07;

import com.sprint.common.excel.data.XCol;
import com.sprint.common.excel.data.XRow;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * ExcelXSSFSheetHandler
 *
 * @author hongfeng-li
 * @version 1.0
 * @since 2020年07月31日
 */
public class ExcelXSSFSheetHandler extends DefaultHandler {

    public static final String ROW = "row";
    public static final String CELL_TYPE_KEY = "t";
    public static final String CELL_STYLE_STR_KEY = "s";
    public static final String ROW_NUM_KEY = "r";
    public static final String CELL_KEY = "c";
    public static final String VAL_KEY = "v";
    public static final String REFERENCE_KEY = "r";

    /**
     * Table with styles
     */
    private final StylesTable stylesTable;

    /**
     * Table with unique strings
     */
    private final ReadOnlySharedStringsTable sharedStringsTable;

    /**
     * Number of columns to read starting with leftmost
     */
    private int minColumnCount = 0;

    /**
     * Set when V start element is seen
     */
    private boolean vIsOpen;

    /**
     * Set when cell start element is seen;
     * used when cell close element is seen.
     */
    private XSSFDataType nextDataType;

    /**
     * Used to format numeric cell values.
     */
    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter = new DataFormatter();

    private int thisColumn = -1;

    /**
     * The last column printed to the output stream
     */
    private int lastColumnNumber = -1;

    private int thisRow = -1;

    /**
     * Gathers characters as they are seen.
     */
    private final StringBuffer value = new StringBuffer();

    private final List<XRow> rows = new ArrayList<XRow>();


    public ExcelXSSFSheetHandler(StylesTable stylesTable, ReadOnlySharedStringsTable sharedStringsTable) {
        this.stylesTable = stylesTable;
        this.sharedStringsTable = sharedStringsTable;
        this.nextDataType = XSSFDataType.NUMBER;
    }

    public List<XRow> getRows() {
        return rows;
    }

    /**
     * (non-Javadoc)
     *
     * @param uri        uri
     * @param localName  localName
     * @param name       name
     * @param attributes attributes
     * @see DefaultHandler#startElement(String, String, String, Attributes)
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        if (XSSFDataType.INLINESTR.getCode().equals(name) || VAL_KEY.equals(name)) {
            vIsOpen = true;
            // Clear contents cache
            value.setLength(0);
        }
        // c => cell
        else if (CELL_KEY.equals(name)) {
            // Get the cell reference
            String r = attributes.getValue(REFERENCE_KEY);
            int firstDigit = -1;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            thisColumn = nameToColumn(r.substring(0, firstDigit));

            // Set up defaults.
            String cellType = attributes.getValue(CELL_TYPE_KEY);
            String cellStyleStr = attributes.getValue(CELL_STYLE_STR_KEY);
            this.formatIndex = -1;
            this.formatString = null;
            this.nextDataType = XSSFDataType.codeOf(cellType, XSSFDataType.NUMBER);
            if (XSSFDataType.NUMBER.equals(nextDataType) && cellStyleStr != null) {
                // It's a number, but almost certainly one
                // with a special style or format
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null) {
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }
        }

        if (ROW.equals(name)) {
            // 容器中加入行
            int rowNum = Integer.parseInt(attributes.getValue(ROW_NUM_KEY));
            rows.add(new XRow(rowNum));
            thisRow++;
        }

    }

    @Override
    public void endElement(String uri, String localName, String name) {

        String thisStr = null;

        // v => contents of a cell
        if (VAL_KEY.equals(name)) {
            // Process the value contents as required.
            // Do now, as characters() may be called more than once
            switch (nextDataType) {

                case BOOL:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;

                case ERROR:
                    thisStr = "ERROR:" + value.toString();
                    break;

                case FORMULA:
                    // A formula could result in a string value,
                    // so always add double-quote characters.
                    thisStr = value.toString();
                    break;

                case INLINESTR:
                    // have seen an example of this, so it's untested.
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    thisStr = rtsi.toString();
                    break;

                case SSTINDEX:
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        RichTextString rtss = sharedStringsTable.getItemAt(idx);
                        thisStr = rtss.toString();
                    } catch (NumberFormatException ignored) {
                    }
                    break;

                case NUMBER:
                    String n = value.toString();
                    if (this.formatString != null) {
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex,
                                this.formatString);
                    } else {
                        thisStr = n;
                    }
                    break;

                default:
                    thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                    break;
            }

            // Output after we've seen the string contents
            // Emit commas for any fields that were missing on this row
            // 测试如果加上这个 第一列数据为null会出现错位
            // if (lastColumnNumber == -1) {
            // lastColumnNumber = 0;
            // }
            int emptyCellNum = thisColumn - lastColumnNumber - 1;
            for (int i = 0; i < emptyCellNum; ++i) {
                rows.get(thisRow).getCol().add(XCol.EMPTY);
            }

            // Might be the empty string.
            XCol col = new XCol(thisStr);
            rows.get(thisRow).getCol().add(col);
            // Update column
            if (thisColumn > -1) {
                lastColumnNumber = thisColumn;
            }

        } else if (ROW.equals(name)) {
            if (thisRow == 0) {
                // 获取标题行的列数
                if (rows.get(0).getCol().size() > minColumnCount) {
                    minColumnCount = rows.get(0).getCol().size();
                }
            }
            // 容器中加入行
            if (minColumnCount > 0) {
                // Columns are 0 based
                if (lastColumnNumber == -1) {
                    lastColumnNumber = 0;
                }
                for (int i = 0; i < (minColumnCount - lastColumnNumber - 1); i++) {
                    rows.get(thisRow).getCol().add(XCol.EMPTY);
                }
            }

            // We're onto a new row
            lastColumnNumber = -1;
        }

    }

    /**
     * Captures characters only if a suitable element is open. Originally was just "v"; extended for inlineStr also.
     */
    @Override
    public void characters(char[] ch, int start, int length) {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
    }

    /**
     * Converts an Excel column name like "C" to a zero-based index.
     *
     * @param name
     * @return Index corresponding to the specified name
     */
    private int nameToColumn(String name) {
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }

}