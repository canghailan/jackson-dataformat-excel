package cc.whohow.excel;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.base.ParserMinimalBase;
import com.fasterxml.jackson.core.io.IOContext;
import com.fasterxml.jackson.core.json.JsonReadContext;
import com.fasterxml.jackson.core.json.PackageVersion;
import com.fasterxml.jackson.core.util.ByteArrayBuilder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayDeque;
import java.util.Deque;
import java.util.List;
import java.util.function.BiFunction;

public class ExcelParser extends ParserMinimalBase {
    private static final int BEFORE_START = -2;
    private static final int START = -1;

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
    protected CellRangeAddress headerRangeAddress;
    protected CellRangeAddress bodyRangeAddress;
    protected List<ColumnKey> keys;

    // parser state
    protected int currentRow;
    protected int currentKey;
    protected Cell currentCell;
    protected boolean eof = false;
    protected boolean closed = false;
    protected Deque<JsonToken> tokenBuffer = new ArrayDeque<>(2);

    public ExcelParser(IOContext ioContext,
                       int features,
                       int excelFeatures,
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
            if (schema.getSheetName() != null) {
                excel = new Excel(workbook.getSheet(schema.getSheetName()));
            } else if (schema.getSheetIndex() >= 0) {
                excel = new Excel(workbook.getSheetAt(schema.getSheetIndex()));
            } else {
                excel = new Excel(workbook.getSheetAt(workbook.getActiveSheetIndex()));
            }
        } catch (InvalidFormatException e) {
            throw new JsonParseException(this, e.getMessage(), e);
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
        excelDetector.detectKeys();
        excelDetector.detectHeaderRangeAddress();
        excelDetector.detectBodyRangeAddress();
        excelDetector.detectKeysIndex();
        keys = excelDetector.getKeys();
        headerRangeAddress = excelDetector.getHeaderRangeAddress();
        bodyRangeAddress = excelDetector.getBodyRangeAddress();

        setCurrentRow(BEFORE_START);
        setCurrentKey(BEFORE_START);
        parsingContext = JsonReadContext.createRootContext(getCurrentRow(), getCurrentColumn(), null);
    }

    protected boolean isStart() {
        return getCurrentRow() < bodyRangeAddress.getFirstRow();
    }

    protected boolean isEnd() {
        return getCurrentRow() > bodyRangeAddress.getLastRow();
    }

    protected boolean isRowStart() {
        return getCurrentKey() <= BEFORE_START;
    }

    protected boolean isRowEnd() {
        return getCurrentKey() >= keys.size() - 1;
    }

    protected void next() throws IOException {
        if (bodyRangeAddress == null) {
            initialize();
        }
        if (isStart()) {
            _handleStart();
            return;
        }
        if (isEnd()) {
            if (eof) {
                _handleEOF();
            } else {
                _handleEnd();
            }
            return;
        }
        if (isRowStart()) {
            _handleRowStart();
            return;
        }
        if (isRowEnd()) {
            _handleRowEnd();
            return;
        }
        _handleCell();
    }

    protected void _handleRowEnd() {
        tokenBuffer.add(JsonToken.END_OBJECT);
        setCurrentRow(getCurrentRow() + 1);
        setCurrentKey(BEFORE_START);
    }

    protected void _handleRowStart() {
        tokenBuffer.add(JsonToken.START_OBJECT);

        CellRangeAddress range = excel.getRowRangeAddress(getCurrentRow());
        range = excel.intersect(range, bodyRangeAddress);
        if (excel.isEmpty(range)) {
            if (skipEmpty) {
                tokenBuffer.clear();
                setCurrentRow(getCurrentRow() + 1);
                return;
            }
        }
        setCurrentKey(START);
    }

    protected void _handleStart() {
        tokenBuffer.add(JsonToken.START_ARRAY);
        setCurrentRow(bodyRangeAddress.getFirstRow());
    }

    protected void _handleEnd() {
        tokenBuffer.add(JsonToken.END_ARRAY);
        eof = true;
    }

    protected void _handleCell() {
        setCurrentKey(getCurrentKey() + 1);
        Cell cell = getCurrentCell();
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
            case START_OBJECT: {
                return getLocation(getCurrentRow(), bodyRangeAddress.getFirstColumn() - 1);
            }
            case END_OBJECT: {
                return getLocation(getCurrentRow() - 1, bodyRangeAddress.getLastColumn() + 1);
            }
            case START_ARRAY: {
                return getLocation(bodyRangeAddress.getFirstRow() - 1, bodyRangeAddress.getFirstColumn() - 1);
            }
            case END_ARRAY: {
                return getLocation(bodyRangeAddress.getLastRow() + 1, bodyRangeAddress.getFirstColumn() - 1);
            }
            default: {
                return getLocation(getCurrentRow(), getCurrentColumn());
            }
        }
    }

    @Override
    public JsonLocation getCurrentLocation() {
        return getLocation(getCurrentRow(), getCurrentColumn());
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
                        getCurrentRow(), bodyRangeAddress.getFirstColumn() - 1);
                break;
            }
            case END_OBJECT: {
                parsingContext = parsingContext.clearAndGetParent();
                break;
            }
            case START_ARRAY: {
                parsingContext = parsingContext.createChildArrayContext(
                        bodyRangeAddress.getFirstRow() - 1, bodyRangeAddress.getFirstColumn() - 1);
                break;
            }
            case END_ARRAY: {
                parsingContext = parsingContext.clearAndGetParent();
                break;
            }
            case FIELD_NAME: {
                parsingContext.setCurrentName(getCurrentColumnKey().getName());
                parsingContext.setCurrentValue(excel.getCellValue(getCurrentCell()));
                break;
            }
        }
        return _currToken;
    }

    @Override
    public JsonParser skipChildren() {
        switch (currentToken()) {
            case START_ARRAY: {
                tokenBuffer.clear();
                setCurrentRow(bodyRangeAddress.getLastRow() + 1);
                setCurrentKey(BEFORE_START);
                return this;
            }
            case START_OBJECT: {
                tokenBuffer.clear();
                setCurrentRow(getCurrentRow() + 1);
                setCurrentKey(BEFORE_START);
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

    @Override
    public String getCurrentName() {
        return parsingContext.getCurrentName();
    }

    protected void setCurrentName(String name) {
        try {
            parsingContext.setCurrentName(name);
        } catch (JsonProcessingException e) {
            throw new UncheckedIOException(e);
        }
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
    public NumberType getNumberType() {
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
        return getCurrentCell().getNumericCellValue();
    }

    @Override
    public BigDecimal getDecimalValue() throws IOException {
        return new BigDecimal(getDoubleValue());
    }

    @Override
    public byte[] getBinaryValue(Base64Variant bv) throws IOException {
        if (currentToken() == JsonToken.VALUE_STRING) {
            Cell cell = getCurrentCell();
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
                return getCurrentCell().getStringCellValue();
            }
            case VALUE_NUMBER_FLOAT:
            case VALUE_NUMBER_INT: {
                return excel.format(getCurrentCell().getNumericCellValue());
            }
            case VALUE_TRUE:
            case VALUE_FALSE: {
                return excel.format(getCurrentCell().getBooleanCellValue());
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

    protected int getCurrentRow() {
        return currentRow;
    }

    protected void setCurrentRow(int row) {
        this.currentRow = row;
        this.currentCell = null;
    }

    protected int getCurrentKey() {
        return currentKey;
    }

    protected void setCurrentKey(int key) {
        this.currentKey = key;
        this.currentCell = null;
    }

    protected int getCurrentColumn() {
        return getColumnByKey(getCurrentKey());
    }

    protected ColumnKey getCurrentColumnKey() {
        return keys.get(getCurrentKey());
    }

    protected int getColumnByKey(int key) {
        if (0 <= key && key < keys.size()) {
            return keys.get(key).getIndex();
        }
        return -1;
    }

    protected Cell getCurrentCell() {
        if (currentCell == null) {
            currentCell = excel.getCell(getCurrentRow(), getCurrentColumn());
        }
        return currentCell;
    }
}
