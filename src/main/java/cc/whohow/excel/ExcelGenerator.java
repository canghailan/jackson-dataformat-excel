package cc.whohow.excel;

import com.fasterxml.jackson.core.Base64Variant;
import com.fasterxml.jackson.core.FormatSchema;
import com.fasterxml.jackson.core.ObjectCodec;
import com.fasterxml.jackson.core.base.GeneratorBase;
import com.fasterxml.jackson.core.json.JsonWriteContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelGenerator extends GeneratorBase {
    // generator props
    protected OutputStream stream;
    protected Workbook template;
    protected Excel excel;
    protected ExcelSchema schema;
    protected CellRangeAddress headerRangeAddress;
    protected CellRangeAddress bodyRangeAddress;

    // generator state
    protected int currentRowIndex;
    protected String currentKeyName;
    protected Row currentRow;
    protected boolean flushed;
    protected Map<String, ColumnKey> keys = new HashMap<>();
    protected Map<Integer, CellStyle> templateCellStyles = new HashMap<>();

    public ExcelGenerator(int features,
                          int excelFeatures,
                          ObjectCodec codec,
                          OutputStream stream) {
        super(features, codec);
        this.stream = stream;
    }

    @Override
    @SuppressWarnings("all")
    public void setSchema(FormatSchema schema) {
        if (schema == null) {
            this.schema = null;
        } else if (schema instanceof ExcelSchema) {
            this.schema = (ExcelSchema) schema;
        } else {
            super.setSchema(schema);
        }
    }

    protected void initialize() throws IOException {
        if (schema == null) {
            schema = new ExcelSchema();
        }
        if (template == null) {
            excel = new Excel(schema.getVersion());
        } else {
            if (schema.getSheetName() != null) {
                excel = new Excel(template.getSheet(schema.getSheetName()));
            } else if (0 <= schema.getSheetIndex() && schema.getSheetIndex() < template.getNumberOfSheets()) {
                excel = new Excel(template.getSheetAt(schema.getSheetIndex()));
            } else if (template.getActiveSheetIndex() < template.getNumberOfSheets()) {
                excel = new Excel(template.getSheetAt(template.getActiveSheetIndex()));
            } else {
                excel = new Excel(template.createSheet());
            }
        }

        if (schema.getHeaderRangeAddress() != null) {
            headerRangeAddress = excel.getCellRangeAddress(schema.getHeaderRangeAddress());
        }
        if (schema.getBodyRangeAddress() != null) {
            bodyRangeAddress = excel.getCellRangeAddress(schema.getBodyRangeAddress());
        }

        ExcelDetector excelDetector = new ExcelDetector(excel);
        excelDetector.setKeys(schema.getKeys());
        excelDetector.setHeaderRangeAddress(headerRangeAddress);
        excelDetector.setBodyRangeAddress(bodyRangeAddress);
        excelDetector.detectHeaderRangeAddress();
        excelDetector.detectBodyRangeAddress();
        excelDetector.detectKeysIndex();
        headerRangeAddress = excelDetector.getHeaderRangeAddress();
        bodyRangeAddress = excelDetector.getBodyRangeAddress();
        List<ColumnKey> keys = excelDetector.getKeys();
        for (ColumnKey key : keys) {
            this.keys.put(key.getName(), key);
        }

        setCurrentRow(-1);
        setCurrentKey(null);

        _writeContext = JsonWriteContext.createRootContext(null);
    }

    @Override
    public void writeStartArray() throws IOException {
        initialize();
        setCurrentRow(bodyRangeAddress.getFirstRow());
    }

    @Override
    public void writeEndArray() throws IOException {
        writeHeader();
    }

    protected void writeHeader() {
        if (headerRangeAddress == null) {
            return;
        }

        Row row = excel.createRow(headerRangeAddress.getLastRow());
        for (ColumnKey key : keys.values()) {
            if (key.getIndex() < 0) {
                continue;
            }
            Cell cell = excel.createCell(row, key.getIndex());
            if (!excel.isEmptyCell(cell)) {
                continue;
            }
            String value = key.getDescription();
            if (value == null || value.isEmpty()) {
                value = key.getName();
            }
            cell.setCellValue(value);
        }
    }

    @Override
    public void writeStartObject() throws IOException {
        createRow();
    }

    @Override
    public void writeEndObject() throws IOException {
        setCurrentRow(getCurrentRow() + 1);
    }

    @Override
    public void writeFieldName(String name) throws IOException {
        _writeContext.writeFieldName(name);
        setCurrentKey(name);
    }

    @Override
    public void writeString(String text) throws IOException {
        _verifyValueWrite("write string");
        setCurrentValue(text);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(text);
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    public void writeString(char[] text, int offset, int length) throws IOException {
        writeString(new String(text, offset, length));
    }

    @Override
    public void writeRawUTF8String(byte[] text, int offset, int length) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeUTF8String(byte[] text, int offset, int length) throws IOException {
        writeString(new String(text, offset, length, StandardCharsets.UTF_8));
    }

    @Override
    public void writeRaw(String text) throws IOException {
        _verifyValueWrite("write raw");
        setCurrentValue(text);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        if (cell.getStringCellValue() == null) {
            cell.setCellValue(text);
            cell.setCellStyle(getTemplateCellStyle(cell));
        } else {
            cell.setCellValue(cell.getStringCellValue() + text);
        }
    }

    @Override
    public void writeRaw(String text, int offset, int len) throws IOException {
        _verifyValueWrite("write raw");
        setCurrentValue(text);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        if (cell.getStringCellValue() == null) {
            cell.setCellValue(text.substring(offset, len));
            cell.setCellStyle(getTemplateCellStyle(cell));
        } else {
            String cellValue = cell.getStringCellValue();
            cell.setCellValue(new StringBuilder(cellValue.length() + len)
                    .append(cellValue)
                    .append(text, offset, len)
                    .toString());
        }
    }

    @Override
    public void writeRaw(char[] text, int offset, int len) throws IOException {
        writeRaw(new String(text, offset, len));
    }

    @Override
    public void writeRaw(char c) throws IOException {
        writeRaw(String.valueOf(c));
    }

    @Override
    public void writeBinary(Base64Variant bv, byte[] data, int offset, int len) throws IOException {
        _verifyValueWrite("write binary");
        byte[] buffer = data;
        if (offset != 0 || len != data.length) {
            buffer = new byte[len];
            System.arraycopy(data, offset, buffer, 0, len);
        }
        String base64 = bv.encode(buffer);
        setCurrentValue(base64);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(base64);
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    public void writeNumber(int v) throws IOException {
        writeNumber((double) v);
    }

    @Override
    public void writeNumber(long v) throws IOException {
        writeNumber((double) v);
    }

    @Override
    public void writeNumber(BigInteger v) throws IOException {
        _verifyValueWrite("write number");
        setCurrentValue(v);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(v.toString());
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    @SuppressWarnings("Duplicates")
    public void writeNumber(double v) throws IOException {
        _verifyValueWrite("write number");
        setCurrentValue(v);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(v);
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    public void writeNumber(float v) throws IOException {
        writeNumber((double) v);
    }

    @Override
    public void writeNumber(BigDecimal v) throws IOException {
        _verifyValueWrite("write number");
        setCurrentValue(v);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(_asString(v));
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    @SuppressWarnings("Duplicates")
    public void writeNumber(String encodedValue) throws IOException {
        _verifyValueWrite("write number");
        setCurrentValue(encodedValue);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(encodedValue);
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    public void writeBoolean(boolean state) throws IOException {
        _verifyValueWrite("write boolean");
        setCurrentValue(state);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellValue(state);
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    public void writeNull() throws IOException {
        _verifyValueWrite("write null");
        setCurrentValue(null);

        Cell cell = createCell();
        if (cell == null) {
            return;
        }
        cell.setCellStyle(getTemplateCellStyle(cell));
    }

    @Override
    public void flush() throws IOException {
        if (flushed) {
            return;
        }
        try {
            excel.getWorkbook().write(stream);
            stream.flush();
        } finally {
            flushed = true;
        }
    }

    @Override
    public void close() throws IOException {
        try {
            flush();
            stream.close();
        } finally {
            super.close();
        }
    }

    @Override
    protected void _releaseBuffers() {
    }

    @Override
    protected void _verifyValueWrite(String typeMsg) throws IOException {
    }

    protected CellStyle getTemplateCellStyle(Cell cell) {
        return templateCellStyles.computeIfAbsent(cell.getColumnIndex(), this::getTemplateCellStyle);
    }

    protected CellStyle getTemplateCellStyle(int column) {
        return excel.createCell(bodyRangeAddress.getFirstRow(), column).getCellStyle();
    }

    protected ColumnKey getColumnKey(String name) {
        return keys.computeIfAbsent(name, this::addColumnKey);
    }

    protected ColumnKey addColumnKey(String name) {
        int index = keys.values().stream()
                .mapToInt(ColumnKey::getIndex)
                .max()
                .orElse(-1);
        return new ColumnKey(name, name, index + 1);
    }

    protected int getCurrentRow() {
        return currentRowIndex;
    }

    protected void setCurrentRow(int row) {
        this.currentRowIndex = row;
    }

    protected void setCurrentKey(String name) {
        this.currentKeyName = name;
    }

    protected ColumnKey getCurrentColumnKey() {
        return getColumnKey(currentKeyName);
    }

    protected Row createRow() {
        if (currentRow == null || currentRow.getRowNum() != currentRowIndex) {
            if (currentRowIndex >= 0) {
                currentRow = excel.createRow(currentRowIndex);
            }
        }
        return currentRow;
    }

    protected Cell createCell() {
        ColumnKey key = getCurrentColumnKey();
        if (key == null || key.getIndex() < 0) {
            return null;
        }
        Row row = createRow();
        if (row == null) {
            return null;
        }
        return excel.createCell(row, key.getIndex());
    }
}
