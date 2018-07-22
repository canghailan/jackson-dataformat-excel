package cc.whohow.excel;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.io.IOContext;

import java.io.*;
import java.net.URL;

public class ExcelFactory extends JsonFactory {
    protected static final String INVALID_FORMAT = "InvalidFormat";
    protected static final int DEFAULT_EXCEL_PARSER_FEATURE_FLAGS = 0;
    protected static final int DEFAULT_EXCEL_GENERATOR_FEATURE_FLAGS = 0;

    public ExcelFactory() {
        super();
    }

    public ExcelFactory(ObjectCodec oc) {
        super(oc);
    }

    protected ExcelFactory(JsonFactory src, ObjectCodec codec) {
        super(src, codec);
    }

    @Override
    public boolean canHandleBinaryNatively() {
        return true;
    }

    @Override
    public boolean canUseCharArrays() {
        return false;
    }

    @Override
    public boolean canUseSchema(FormatSchema schema) {
        return schema == null || schema instanceof ExcelSchema;
    }

    @Override
    public String getFormatName() {
        return "EXCEL";
    }

    @Override
    public ExcelParser createParser(File f) throws IOException, JsonParseException {
        return createParser(new FileInputStream(f));
    }

    @Override
    public ExcelParser createParser(URL url) throws IOException, JsonParseException {
        return createParser(_optimizedStreamFromURL(url));
    }

    @Override
    public ExcelParser createParser(InputStream in) throws IOException, JsonParseException {
        return _createParser(in, null);
    }

    @Override
    public ExcelParser createParser(Reader r) throws IOException, JsonParseException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelParser createParser(byte[] data) throws IOException, JsonParseException {
        return createParser(data, 0, data.length);
    }

    @Override
    public ExcelParser createParser(byte[] data, int offset, int len) throws IOException, JsonParseException {
        return createParser(new ByteArrayInputStream(data, offset, len));
    }

    @Override
    public ExcelParser createParser(String content) throws IOException, JsonParseException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelParser createParser(char[] content) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelParser createParser(char[] content, int offset, int len) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelParser createParser(DataInput in) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelGenerator createGenerator(OutputStream out, JsonEncoding enc) throws IOException {
        return createGenerator(out);
    }

    @Override
    public ExcelGenerator createGenerator(OutputStream out) throws IOException {
        return _createUTF8Generator(out, null);
    }

    @Override
    public ExcelGenerator createGenerator(Writer w) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelGenerator createGenerator(File f, JsonEncoding enc) throws IOException {
        return createGenerator(new FileOutputStream(f), enc);
    }

    @Override
    public ExcelGenerator createGenerator(DataOutput out, JsonEncoding enc) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    public ExcelGenerator createGenerator(DataOutput out) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    protected ExcelParser _createParser(InputStream in, IOContext ioContext) throws IOException {
        return new ExcelParser(ioContext, _parserFeatures, DEFAULT_EXCEL_PARSER_FEATURE_FLAGS, _objectCodec, in);
    }

    @Override
    protected ExcelParser _createParser(Reader r, IOContext ioContext) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    protected ExcelParser _createParser(char[] data, int offset, int len, IOContext ioContext, boolean recyclable) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    protected ExcelParser _createParser(byte[] data, int offset, int len, IOContext ioContext) throws IOException {
        return _createParser(new ByteArrayInputStream(data, offset, len), ioContext);
    }

    @Override
    protected ExcelParser _createParser(DataInput input, IOContext ioContext) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    protected ExcelGenerator _createGenerator(Writer out, IOContext ioContext) throws IOException {
        throw new UnsupportedOperationException(INVALID_FORMAT);
    }

    @Override
    protected ExcelGenerator _createUTF8Generator(OutputStream out, IOContext ioContext) throws IOException {
        return new ExcelGenerator(_generatorFeatures, DEFAULT_EXCEL_GENERATOR_FEATURE_FLAGS, _objectCodec, out);
    }
}
