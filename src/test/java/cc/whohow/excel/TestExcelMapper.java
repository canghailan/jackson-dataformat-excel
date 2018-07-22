package cc.whohow.excel;

import org.junit.Test;

import java.io.File;

public class TestExcelMapper {
    @Test
    public void test() throws Exception {
        ExcelMapper mapper = new ExcelMapper();
        System.out.println(mapper.readTree(new File("test.xlsx")));
    }
}
