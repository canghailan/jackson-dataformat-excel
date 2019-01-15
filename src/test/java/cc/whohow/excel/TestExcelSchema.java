package cc.whohow.excel;

import org.junit.Test;

public class TestExcelSchema {
    @Test
    public void test() {
        ExcelMapper objectMapper = new ExcelMapper();
        System.out.println(objectMapper.schemaForReader(DataModel1.class).getKeys());
        System.out.println(objectMapper.schemaForWriter(DataModel2.class).getKeys());
    }
}
