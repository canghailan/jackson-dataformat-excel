package cc.whohow.excel;

import com.fasterxml.jackson.core.type.TypeReference;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class TestExcelMapper {
    @Test
    public void test1() throws Exception {
        ExcelMapper excelMapper = new ExcelMapper();
        System.out.println(excelMapper.readTree(new File("test.xlsx")));
        System.out.println(excelMapper.readValue(new File("test.xlsx"), new TypeReference<List<Map<String, String>>>() {
        }).toString());
        System.out.println(excelMapper.readValue(new File("test.xlsx"), new TypeReference<List<DataModel1>>() {
        }).toString());
    }

    @Test
    public void test2() throws Exception {
        ExcelMapper excelMapper = new ExcelMapper();
        System.out.println(excelMapper.readTree(new File("test2.xls")));
        System.out.println(excelMapper.readValue(new File("test2.xls"), new TypeReference<List<Map<String, String>>>() {
        }).toString());
        System.out.println(excelMapper.readValue(new File("test2.xls"), new TypeReference<List<DataModel2>>() {
        }).toString());
    }
}
