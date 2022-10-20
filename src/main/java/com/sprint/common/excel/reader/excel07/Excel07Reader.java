package com.sprint.common.excel.reader.excel07;

import com.sprint.common.excel.data.Sheet;
import com.sprint.common.excel.data.XRow;
import com.sprint.common.excel.reader.ExcelReader;
import com.sprint.common.excel.util.Closeables;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

public class Excel07Reader implements ExcelReader {

    private OPCPackage xlsxPackage;

    private Map<String, List<XRow>> allContainer = new LinkedHashMap<>();

    private Function<String, Boolean> sheetFilter;

    public Excel07Reader(OPCPackage pkg) {
        this(pkg, (sheetName) -> true);
    }

    public Excel07Reader(OPCPackage pkg, Function<String, Boolean> sheetFilter) {
        this.xlsxPackage = pkg;
        this.sheetFilter = sheetFilter;
    }


    /**
     * 处理sheet
     *
     * @param sheetName        sheetName
     * @param styles           styles
     * @param strings          strings
     * @param sheetInputStream sheetInputStream
     * @throws IOException                  IOException
     * @throws ParserConfigurationException ParserConfigurationException
     * @throws SAXException                 SAXException
     */
    public void processSheet(String sheetName, StylesTable styles, ReadOnlySharedStringsTable strings,
                             InputStream sheetInputStream) throws IOException, ParserConfigurationException, SAXException {
        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();
        ExcelXSSFSheetHandler handler = new ExcelXSSFSheetHandler(styles, strings);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
        allContainer.computeIfAbsent(sheetName, (key) -> new ArrayList<>()).addAll(handler.getRows());
    }

    @Override
    public List<Sheet> process() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (iter.hasNext()) {
            InputStream stream = null;
            try {
                stream = iter.next();
                String name = iter.getSheetName();
                if (sheetFilter.apply(name)) {
                    processSheet(name, styles, strings, stream);
                }
            } finally {
                Closeables.close(stream);
            }
        }
        return getSheets();
    }

    private List<Sheet> getSheets() {
        int sheetNumber = 1;
        List<Sheet> sheets = new ArrayList<>(allContainer.size());
        for (Map.Entry<String, List<XRow>> entry : allContainer.entrySet()) {
            sheets.add(new Sheet(sheetNumber++, entry.getKey(), entry.getValue()));
        }
        return sheets;
    }
}