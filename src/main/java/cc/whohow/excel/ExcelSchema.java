package cc.whohow.excel;

import com.fasterxml.jackson.core.FormatSchema;
import org.apache.poi.ss.SpreadsheetVersion;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class ExcelSchema implements FormatSchema {
    private SpreadsheetVersion version = SpreadsheetVersion.EXCEL2007;
    private int sheetIndex = -1;
    private String sheetName = null;
    private String headerRangeAddress = "1:1";
    private String dataRangeAddress = "2:";
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

    public ExcelSchema withData(String data) {
        this.dataRangeAddress = data;
        return this;
    }

    public ExcelSchema addKey(String name) {
        return addKey(name, null);
    }

    public ExcelSchema addKey(String name, String description) {
        this.keys.add(new ColumnKey(name, description));
        return this;
    }

    public ExcelSchema addKey(int index, String name, String description) {
        while (keys.size() <= index) {
            keys.add(null);
        }
        this.keys.set(index, new ColumnKey(name, description));
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

    public String getDataRangeAddress() {
        return dataRangeAddress;
    }

    public List<ColumnKey> getKeys() {
        keys.removeIf(Objects::isNull);
        return keys.isEmpty() ? null : keys;
    }
}
