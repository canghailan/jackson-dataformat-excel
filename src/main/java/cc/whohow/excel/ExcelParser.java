package cc.whohow.excel;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.base.ParserMinimalBase;
import com.fasterxml.jackson.core.io.IOContext;
import com.fasterxml.jackson.core.json.JsonReadContext;
import com.fasterxml.jackson.core.json.PackageVersion;
import com.fasterxml.jackson.core.util.ByteArrayBuilder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.*;
import java.util.function.BiFunction;

public class ExcelParser extends ParserMinimalBase {
    // parser context
    protected ObjectCodec codec;
    protected IOContext ioContext;
    protected JsonReadContext parsingContext;

    // parser feature
    protected boolean skipEmpty = true;

    // parser props
    protected InputStream stream;
    protected Excel excel;
    protected ExcelSchema schema;
    protected CellRangeAddress dataRangeAddress;
    protected List<ColumnKey> keys;

    // parser state
    protected int r;
    protected int c;
    protected Row row;
    protected Cell cell;
    protected boolean eof = false;
    protected boolean closed = false;
    protected Deque<JsonToken> tokenBuffer = new ArrayDeque<>(2);

    public ExcelParser(IOContext ioContext,
                       int features, int excelFeatures,
                       ObjectCodec codec,
                       InputStream stream) {
        super(features);
        this.ioContext = ioContext;
        this.codec = codec;
        this.stream = stream;
    }

    public ExcelParser(int features,
                       InputStream stream) {
        this(null, features, 0, null, stream);
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
        try {
            Workbook workbook = WorkbookFactory.create(stream);
            Sheet sheet;
            if (schema.getSheetName() != null) {
                sheet = workbook.getSheet(schema.getSheetName());
            } else {
                int sheetIndex = schema.getSheetIndex();
                if (sheetIndex < 0) {
                    sheetIndex = workbook.getActiveSheetIndex();
                }
                sheet = workbook.getSheetAt(sheetIndex);
            }
            excel = new Excel(sheet);
        } catch (InvalidFormatException e) {
            throw new JsonParseException(this, e.getMessage(), e);
        }

        keys = schema.getKeys();
        if (keys == null) {
            keys = getKeysFromHeader(excel.getCellRangeAddress(schema.getHeaderRangeAddress()));
        }
        if (keys == null) {
            keys = new ExcelColumnKeys();
        }

        dataRangeAddress = excel.getCellRangeAddress(schema.getDataRangeAddress());

        row = null;
        cell = null;

        r = dataRangeAddress.getFirstRow() - 1;
        c = dataRangeAddress.getFirstColumn() - 1;

        parsingContext = JsonReadContext.createRootContext(r, c, null);
    }

    protected void next() throws IOException {
        if (dataRangeAddress == null) {
            initialize();
        }
        if (r < dataRangeAddress.getFirstRow()) {
            _handleStart();
            return;
        }
        if (r > dataRangeAddress.getLastRow()) {
            if (c < dataRangeAddress.getFirstColumn()) {
                _handleEnd();
            } else {
                _handleEOF();
            }
            return;
        }
        if (c < dataRangeAddress.getFirstColumn()) {
            _handleRowStart();
            return;
        }
        if (c > dataRangeAddress.getLastColumn() || row == null || c >= row.getLastCellNum()) {
            _handleRowEnd();
            return;
        }
        _handleCell();
    }

    protected void _handleRowEnd() {
        tokenBuffer.add(JsonToken.END_OBJECT);
        r++;
        c = dataRangeAddress.getFirstColumn() - 1;
        cell = null;
    }

    protected void _handleRowStart() {
        tokenBuffer.add(JsonToken.START_OBJECT);
        row = excel.getRow(r);

        if (excel.isEmptyRow(row, dataRangeAddress)) {
            if (skipEmpty) {
                tokenBuffer.clear();
                r++;
                c = dataRangeAddress.getFirstColumn() - 1;
                cell = null;
                return;
            }
        }

        c = dataRangeAddress.getFirstColumn();
        if (row != null && c < row.getFirstCellNum()) {
            c = row.getFirstCellNum();
        }
        cell = null;
    }

    protected void _handleStart() {
        tokenBuffer.add(JsonToken.START_ARRAY);
        r = dataRangeAddress.getFirstRow();
        c = dataRangeAddress.getFirstColumn() - 1;
        cell = null;
    }

    protected void _handleEnd() {
        tokenBuffer.add(JsonToken.END_ARRAY);
        r = dataRangeAddress.getLastRow() + 1;
        c = dataRangeAddress.getFirstColumn();
        cell = null;
    }

    protected void _handleCell() {
        cell = excel.getCell(r, c++);
        if (cell == null) {
            return;
        }
        switch (cell.getCellTypeEnum()) {
            case STRING: {
                tokenBuffer.add(JsonToken.FIELD_NAME);
                tokenBuffer.add(JsonToken.VALUE_STRING);
                break;
            }
            case NUMERIC: {
                tokenBuffer.add(JsonToken.FIELD_NAME);
                tokenBuffer.add(JsonToken.VALUE_NUMBER_FLOAT);
                break;
            }
            case BOOLEAN: {
                tokenBuffer.add(JsonToken.FIELD_NAME);
                if (cell.getBooleanCellValue()) {
                    tokenBuffer.add(JsonToken.VALUE_TRUE);
                } else {
                    tokenBuffer.add(JsonToken.VALUE_FALSE);
                }
                break;
            }
            default: {
                tokenBuffer.add(JsonToken.FIELD_NAME);
                tokenBuffer.add(JsonToken.VALUE_NULL);
                break;
            }
        }
    }

    @Override
    protected void _handleEOF() {
        eof = true;
    }

    @Override
    public ObjectCodec getCodec() {
        return codec;
    }

    @Override
    public void setCodec(ObjectCodec c) {
        this.codec = c;
    }

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    @Override
    public void close() throws IOException {
        try {
            stream.close();
        } finally {
            closed = true;
        }
    }

    @Override
    public boolean isClosed() {
        return closed;
    }

    @Override
    public JsonStreamContext getParsingContext() {
        return parsingContext;
    }

    @Override
    public JsonLocation getTokenLocation() {
        JsonToken token = currentToken();
        if (token == null) {
            return getCurrentLocation();
        }
        switch (token) {
            case START_OBJECT:{
                return getLocation(r, dataRangeAddress.getFirstColumn() - 1);
            }
            case END_OBJECT: {
                return getLocation(r - 1, dataRangeAddress.getLastColumn() + 1);
            }
            case START_ARRAY:{
                return getLocation(dataRangeAddress.getFirstRow() - 1, dataRangeAddress.getFirstColumn() - 1);
            }
            case END_ARRAY: {
                return getLocation(dataRangeAddress.getLastRow() + 1, dataRangeAddress.getFirstColumn() - 1);
            }
            default: {
                return getLocation(r, c - 1);
            }
        }
    }

    @Override
    public JsonLocation getCurrentLocation() {
        return getLocation(r, c);
    }

    protected JsonLocation getLocation(int row, int column) {
        return new JsonLocation(excel.getSheet(), -1L, row, column);
    }

    @Override
    public JsonToken nextToken() throws IOException {
        while (tokenBuffer.isEmpty() && !eof) {
            next();
        }
        _currToken = tokenBuffer.pollFirst();
        if (_currToken == null) {
            return null;
        }
        switch (_currToken) {
            case START_OBJECT: {
                parsingContext = parsingContext.createChildObjectContext(
                        r, dataRangeAddress.getFirstColumn() - 1);
                break;
            }
            case END_OBJECT: {
                parsingContext = parsingContext.clearAndGetParent();
                break;
            }
            case START_ARRAY: {
                parsingContext = parsingContext.createChildArrayContext(
                        dataRangeAddress.getFirstRow() - 1, dataRangeAddress.getFirstColumn() - 1);
                break;
            }
            case END_ARRAY: {
                parsingContext = parsingContext.clearAndGetParent();
                break;
            }
            case FIELD_NAME: {
                parsingContext.setCurrentName(getCellName(cell));
                parsingContext.setCurrentValue(excel.getCellValue(cell));
                break;
            }
        }
        return _currToken;
    }

    @Override
    public JsonParser skipChildren() {
        switch (currentToken()) {
            case START_ARRAY: {
                r = dataRangeAddress.getLastRow() + 1;
                c = dataRangeAddress.getFirstColumn() - 1;

                row = null;
                cell = null;

                tokenBuffer.clear();
                return this;
            }
            case START_OBJECT: {
                c = dataRangeAddress.getLastColumn() + 1;

                cell = null;

                tokenBuffer.clear();
                return this;
            }
            default: {
                return this;
            }
        }
    }

    @Override
    public void overrideCurrentName(String name) {
        setCurrentName(name);
    }

    protected void setCurrentName(String name) {
        try {
            parsingContext.setCurrentName(name);
        } catch (JsonProcessingException e) {
            throw new UncheckedIOException(e);
        }
    }

    @Override
    public String getCurrentName() {
        return parsingContext.getCurrentName();
    }

    @Override
    public String getText() {
        return getCurrentTokenAsString(null, this::asString);
    }

    @Override
    public char[] getTextCharacters() {
        String text = getText();
        return text == null ? null : text.toCharArray();
    }

    @Override
    public int getTextLength() {
        String text = getText();
        return text == null ? 0 : text.length();
    }

    @Override
    public int getTextOffset() {
        return 0;
    }

    @Override
    public boolean hasTextCharacters() {
        return false;
    }

    @Override
    public Number getNumberValue() throws IOException {
        return getDoubleValue();
    }

    @Override
    public NumberType getNumberType()  {
        JsonToken token = currentToken();
        if (token == JsonToken.VALUE_NUMBER_FLOAT || token == JsonToken.VALUE_NUMBER_INT) {
            return NumberType.DOUBLE;
        }
        return null;
    }

    @Override
    public int getIntValue() throws IOException {
        double value = getDoubleValue();
        if (Integer.MIN_VALUE <= value && value <= Integer.MAX_VALUE) {
            return (int) value;
        }
        reportOverflowInt();
        return 0;
    }

    @Override
    public long getLongValue() throws IOException {
        double value = getDoubleValue();
        if (Long.MIN_VALUE <= value && value <= Long.MAX_VALUE) {
            return (int) value;
        }
        reportOverflowLong();
        return 0L;
    }

    @Override
    public BigInteger getBigIntegerValue() throws IOException {
        return getDecimalValue().toBigInteger();
    }

    @Override
    public float getFloatValue() throws IOException {
        return (float) getDoubleValue();
    }

    @Override
    public double getDoubleValue() throws IOException {
        return cell.getNumericCellValue();
    }

    @Override
    public BigDecimal getDecimalValue() throws IOException {
        return new BigDecimal(getDoubleValue());
    }

    @Override
    public byte[] getBinaryValue(Base64Variant bv) throws IOException {
        if (currentToken() == JsonToken.VALUE_STRING) {
            String base64 = cell.getStringCellValue();
            if (base64 == null || base64.isEmpty()) {
                return NO_BYTES;
            }
            ByteArrayBuilder builder = new ByteArrayBuilder(base64.length() / 4 * 3 + 1);
            try {
                _decodeBase64(cell.getStringCellValue(), builder, bv);
                return builder.toByteArray();
            } finally {
                builder.release();
                builder.close();
            }
        }
        _reportError("Current token (%s) not VALUE_STRING, can not access as binary", _currToken);
        return null;
    }

    @Override
    public String getValueAsString(String defaultValue) throws IOException {
        return getCurrentTokenAsString(defaultValue, this::defaultValue);
    }

    protected String getCurrentTokenAsString(String defaultValue,
                                             BiFunction<JsonToken, String, String> notValueTokenMapper) {
        JsonToken token = currentToken();
        if (token == null) {
            return defaultValue;
        }
        switch (token) {
            case FIELD_NAME: {
                return getCurrentName();
            }
            case VALUE_STRING: {
                return cell.getStringCellValue();
            }
            case VALUE_NUMBER_FLOAT:
            case VALUE_NUMBER_INT:{
                return excel.format(cell.getNumericCellValue());
            }
            case VALUE_TRUE:
            case VALUE_FALSE: {
                return excel.format(cell.getBooleanCellValue());
            }
            case VALUE_NULL: {
                return defaultValue;
            }
            default: {
                return notValueTokenMapper.apply(token, defaultValue);
            }
        }
    }

    private String asString(JsonToken token, String defaultValue) {
        return token.asString();
    }

    private String defaultValue(JsonToken token, String defaultValue) {
        return defaultValue;
    }

    protected String getCellName(Cell cell) {
        if (cell == null) {
            return null;
        }
        return keys.get(cell.getColumnIndex() - dataRangeAddress.getFirstColumn()).getName();
    }


    protected List<ColumnKey> getKeysFromHeader(CellRangeAddress header) {
        if (header == null) {
            return null;
        }
        List<ColumnKey> keys = new ArrayList<>();
        for (int c = header.getFirstColumn(), i = 0; c <= header.getLastColumn(); c++, i++) {
            Cell cell = excel.getCell(header.getLastRow(), c);
            String text = excel.formatCellValue(cell, null);
            if (text == null || text.isEmpty()) {
                return keys;
            }
            keys.add(new ColumnKey(text, text));
        }
        return keys.isEmpty() ? null : keys;
    }
}
