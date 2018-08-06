package cc.whohow.excel;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JavaType;
import org.junit.Test;

import java.util.List;

public class TestExcelSchema {
    @Test
    public void test() {
        ExcelMapper objectMapper = new ExcelMapper();
        JavaType type = objectMapper.getTypeFactory().constructType(new TypeReference<List<DataModel>>() {
        });
        System.out.println(objectMapper.schemaFor(type, objectMapper.getDeserializationConfig()::introspect).getKeys());
    }
}
