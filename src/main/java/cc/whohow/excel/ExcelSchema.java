package cc.whohow.excel;

import com.fasterxml.jackson.core.FormatSchema;
import org.apache.poi.ss.SpreadsheetVersion;

import java.util.ArrayList;
import java.util.List;

public class ExcelSchema implements FormatSchema {
    private SpreadsheetVersion version = SpreadsheetVersion.EXCEL2007;
    private int sheetIndex = -1;
    private String sheetName = null;
    private String headerRangeAddress = null;
    private String bodyRangeAddress = null;
    private List<ColumnKey> keys = new ArrayList<>();

    @Override
    public String getSchemaType() {
        return "EXCEL";
    }

    public ExcelSchema withVersion(String version) {
        return withVersion(SpreadsheetVersion.valueOf(version));
    }

    public ExcelSchema withVersion(SpreadsheetVersion version) {
        this.version = version;
        return this;
    }

    public ExcelSchema withSheet(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    public ExcelSchema withSheet(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelSchema withHeader(String header) {
        this.headerRangeAddress = header;
        return this;
    }

    public ExcelSchema withBody(String body) {
        this.bodyRangeAddress = body;
        return this;
    }

    public ExcelSchema withKeys(List<ColumnKey> keys) {
        this.keys.clear();
        this.keys.addAll(keys);
        return this;
    }

    public ExcelSchema addKey(String name) {
        this.keys.add(new ColumnKey(name));
        return this;
    }

    public ExcelSchema addKey(String name, String description) {
        this.keys.add(new ColumnKey(name, description));
        return this;
    }

    public ExcelSchema addKey(String name, String description, int index) {
        this.keys.add(new ColumnKey(name, description, index));
        return this;
    }

    public SpreadsheetVersion getVersion() {
        return version;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getHeaderRangeAddress() {
        return headerRangeAddress;
    }

    public String getBodyRangeAddress() {
        return bodyRangeAddress;
    }

    public List<ColumnKey> getKeys() {
        return keys;
    }

    public void detect(Excel excel) {
        ExcelDetector excelDetector = new ExcelDetector(excel);
        excelDetector.withKeys(keys);
        if (headerRangeAddress != null) {
            excelDetector.withHeaderRangeAddress(excel.getCellRangeAddress(headerRangeAddress));
        }
        if (bodyRangeAddress != null) {
            excelDetector.withBodyRangeAddress(excel.getCellRangeAddress(bodyRangeAddress));
        }
        if (excelDetector.call()) {
            keys = excelDetector.getKeys();
            headerRangeAddress = excelDetector.getHeaderRangeAddress().formatAsString();
            bodyRangeAddress = excelDetector.getBodyRangeAddress().formatAsString();
        }
    }
}
