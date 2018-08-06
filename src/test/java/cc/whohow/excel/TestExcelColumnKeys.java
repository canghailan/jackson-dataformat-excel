package cc.whohow.excel;

import org.junit.Test;

public class TestExcelColumnKeys {
    @Test
    public void test() {
        ExcelColumnKeys names = new ExcelColumnKeys();
        names.subList(0, 100).forEach(System.out::println);
    }
}
