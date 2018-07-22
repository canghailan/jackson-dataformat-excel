package cc.whohow.excel;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.File;
import java.io.IOException;

public class ExcelMapper extends ObjectMapper {
    public ExcelMapper() {
        this(new ExcelFactory());
    }

    public ExcelMapper(ExcelFactory factory) {
        super(factory);
    }

    public ExcelMapper(ExcelMapper mapper) {
        super(mapper);
    }

    @Override
    public ExcelMapper copy() {
        _checkInvalidCopy(ExcelMapper.class);
        return new ExcelMapper(this);
    }

    @Override
    public ExcelFactory getFactory() {
        return (ExcelFactory) _jsonFactory;
    }

    public String[][] readText(File file) throws IOException {
        return null;
    }
}
