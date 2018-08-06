package cc.whohow.excel;

import com.fasterxml.jackson.core.type.TypeReference;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class TestExcelMapper {
    @Test
    public void test() throws Exception {
        ExcelMapper mapper = new ExcelMapper();
        System.out.println(mapper.readTree(new File("test.xlsx")));
        System.out.println(mapper.readValue(new File("test.xlsx"), new TypeReference<List<Map<String, String>>>() {
        }).toString());
        System.out.println(mapper.readValue(new File("test.xlsx"), new TypeReference<List<DataModel>>() {
        }).toString());
    }
}
