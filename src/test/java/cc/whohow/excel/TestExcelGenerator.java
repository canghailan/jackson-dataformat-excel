package cc.whohow.excel;

import com.fasterxml.jackson.databind.JsonNode;
import org.junit.Test;

import java.io.File;

public class TestExcelGenerator {
    @Test
    public void test() throws Exception {
        ExcelMapper mapper = new ExcelMapper();
        JsonNode data = mapper.readTree(new File("test.xlsx"));
        System.out.println(data);
        mapper.writeValue(new File("test-generator.xlsx"), data);
    }
}
