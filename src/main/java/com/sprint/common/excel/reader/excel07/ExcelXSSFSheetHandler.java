package com.sprint.common.excel.reader.excel07;

import com.sprint.common.excel.data.XCol;
import com.sprint.common.excel.data.XRow;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
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

    /**
     * The type of the data value is indicated by an attribute on the cell. The value is usually in a "v" element within
     * the cell.
     */
    enum XSSFDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
    }

    /**
     * Table with styles
     */
    private StylesTable stylesTable;

    /**
     * Table with unique strings
     */
    private ReadOnlySharedStringsTable sharedStringsTable;

    /**
     * Number of columns to read starting with leftmost
     */
    private int minColumnCount = 0;

    // Set when V start element is seen
    private boolean vIsOpen;

    // Set when cell start element is seen;
    // used when cell close element is seen.
    private XSSFDataType nextDataType;

    // Used to format numeric cell values.
    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter;

    private int thisColumn = -1;
    // The last column printed to the output stream
    private int lastColumnNumber = -1;

    // Gathers characters as they are seen.
    private StringBuffer value;

    private List<XRow> rows = new ArrayList<XRow>();

    private int thisRow = -1;

    /**
     * Accepts objects needed while parsing.
     *
     * @param styles  Table of styles
     * @param strings Table of shared strings
     */
    public ExcelXSSFSheetHandler(StylesTable styles, ReadOnlySharedStringsTable strings) {
        this.stylesTable = styles;
        this.sharedStringsTable = strings;
        this.value = new StringBuffer();
        this.nextDataType = XSSFDataType.NUMBER;
        this.formatter = new DataFormatter();
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
        if ("inlineStr".equals(name) || "v".equals(name)) {
            vIsOpen = true;
            // Clear contents cache
            value.setLength(0);
        }
        // c => cell
        else if ("c".equals(name)) {
            // Get the cell reference
            String r = attributes.getValue("r");
            int firstDigit = -1;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            thisColumn = nameToColumn(r.substring(0, firstDigit));

            // Set up defaults.
            this.nextDataType = XSSFDataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType)) {
                nextDataType = XSSFDataType.BOOL;
            } else if ("e".equals(cellType)) {
                nextDataType = XSSFDataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                nextDataType = XSSFDataType.INLINESTR;
            } else if ("s".equals(cellType)) {
                nextDataType = XSSFDataType.SSTINDEX;
            } else if ("str".equals(cellType)) {
                nextDataType = XSSFDataType.FORMULA;
            } else if (cellStyleStr != null) {
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

        if ("row".equals(name)) {
            // 容器中加入行
            int rowNum = Integer.parseInt(attributes.getValue("r"));
            rows.add(new XRow(rowNum));
            thisRow++;
        }

    }

    @Override
    public void endElement(String uri, String localName, String name) {

        String thisStr = null;

        // v => contents of a cell
        if ("v".equals(name)) {
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
                        XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                        thisStr = rtss.toString();
                    } catch (NumberFormatException ex) {
                        // output.println("Failed to parse SST index '" + sstIndex + "': " + ex.toString());
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

        } else if ("row".equals(name)) {
            if (thisRow == 0) {
                // 获取标题行的列数
                if (rows.get(0).getCol().size() > minColumnCount) {
                    minColumnCount = rows.get(0).getCol().size();
                }
            }
            // 容器中加入行
            // rows.add(new XRow());
            // thisRow++;
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