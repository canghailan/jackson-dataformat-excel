package cc.whohow.excel;

import org.junit.Test;

public class TestExcelColumnKeys {
    @Test
    public void test() {
        ExcelColumnKeys names = new ExcelColumnKeys();
        names.forEach(System.out::println);
    }
}
