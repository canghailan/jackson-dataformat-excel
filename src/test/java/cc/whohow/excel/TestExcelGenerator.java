package cc.whohow.excel;

import com.fasterxml.jackson.core.type.TypeReference;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class TestExcelGenerator {
    @Test
    public void test() throws Exception {
        ExcelMapper mapper = new ExcelMapper();
        ExcelSchema schema = mapper.schemaForReader(DataModel1.class);
        List<DataModel1> data = mapper.reader(schema).forType(new TypeReference<List<DataModel1>>() {
        }).readValue(new File("test.xlsx"));
        System.out.println(data);
        mapper.writer(schema).writeValue(new File("test-generator.xlsx"), data);
    }
}
