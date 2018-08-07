package cc.whohow.excel;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.base.GeneratorBase;
import com.fasterxml.jackson.core.json.JsonWriteContext;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class ExcelGenerator extends GeneratorBase {
    // generator props
    protected OutputStream stream;
    protected Workbook template;
    protected Excel excel;
    protected ExcelSchema schema;
    protected CellRangeAddress headerRangeAddress;
    protected CellRangeAddress bodyRangeAddress;
    protected List<ColumnKey> keys;

    // generator state
    protected int r;
    protected int c;
    protected Row row;
    protected Cell cell;
    protected boolean flushed;
    protected Map<String, ColumnKey> keyIndex = new HashMap<>();
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
            } else if (0 <= template.getActiveSheetIndex() && template.getActiveSheetIndex() < template.getNumberOfSheets()) {
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
        keys = excelDetector.getKeys();
        headerRangeAddress = excelDetector.getHeaderRangeAddress();
        bodyRangeAddress = excelDetector.getBodyRangeAddress();

        row = null;
        cell = null;

        r = bodyRangeAddress.getFirstRow() - 1;
        c = bodyRangeAddress.getFirstColumn() - 1;

        _writeContext = JsonWriteContext.createRootContext(null);
    }

    @Override
    public void writeStartArray() throws IOException {
        initialize();

        r = bodyRangeAddress.getFirstRow();
        c = bodyRangeAddress.getFirstColumn() - 1;
        row = null;
        cell = null;
    }

    @Override
    public void writeEndArray() throws IOException {
        writeHeader();

        c = bodyRangeAddress.getFirstColumn() - 1;
        row = null;
        cell = null;
    }

    protected void writeHeader() {
        if (headerRangeAddress == null) {
            return;
        }

        r = headerRangeAddress.getLastRow();
        row = excel.createRow(r);

        c = headerRangeAddress.getFirstColumn();
        for (ColumnKey key : keys) {
            cell = excel.createCell(row, c);

            if (key.getDescription() == null || key.getDescription().isEmpty()) {
                cell.setCellValue(key.getName());
            } else {
                cell.setCellValue(key.getDescription());
            }

            c++;
        }
    }

    @Override
    public void writeStartObject() throws IOException {
        c = bodyRangeAddress.getFirstColumn();
        row = excel.createRow(r);
        cell = null;
    }

    @Override
    public void writeEndObject() throws IOException {
        r++;
        c = bodyRangeAddress.getFirstColumn() - 1;
        cell = null;
    }

    @Override
    public void writeFieldName(String name) throws IOException {
        _writeContext.writeFieldName(name);

        c = getColumnKey(name).getIndex();
        if (c < 0) {
            return;
        }
        cell = excel.createCell(row, c);
    }

    @Override
    public void writeString(String text) throws IOException {
        _verifyValueWrite("write string");
        setCurrentValue(text);

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

    protected ColumnKey getColumnKey(String name) {
        return keyIndex.computeIfAbsent(name, this::findOrAddColumnKey);
    }

    protected ColumnKey findColumnKey(String name) {
        for (ColumnKey key : keys) {
            if (key.getName().equals(name)) {
                return key;
            }
        }
        return null;
    }

    protected ColumnKey addColumnKey(String name) {
        int index = keys.stream()
                .mapToInt(ColumnKey::getIndex)
                .max()
                .orElse(-1);
        return new ColumnKey(name, name, index + 1);
    }

    protected ColumnKey findOrAddColumnKey(String name) {
        ColumnKey key = findColumnKey(name);
        if (key == null) {
            key = addColumnKey(name);
        }
        return key;
    }

    protected CellStyle getTemplateCellStyle(Cell cell) {
        return templateCellStyles.computeIfAbsent(cell.getColumnIndex(), this::getTemplateCellStyle);
    }

    protected CellStyle getTemplateCellStyle(int column) {
        return excel.createCell(bodyRangeAddress.getFirstRow(), column).getCellStyle();
    }
}
