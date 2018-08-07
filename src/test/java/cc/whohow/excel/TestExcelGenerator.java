package cc.whohow.excel;

import com.fasterxml.jackson.core.type.TypeReference;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class TestExcelGenerator {
    @Test
    public void test() throws Exception {
        ExcelMapper mapper = new ExcelMapper();
        List<DataModel> data = mapper.readValue(new File("test.xlsx"), new TypeReference<List<DataModel>>() {
        });
        System.out.println(data);
        mapper.writerFor(new TypeReference<List<DataModel>>() {
        }).writeValue(new File("test-generator.xlsx"), data);
    }
}
