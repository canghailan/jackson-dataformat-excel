package cc.whohow.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.junit.Assert;
import org.junit.Test;

public class TestReference {
    @Test
    public void test() {
        System.out.println(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
        System.out.println(new CellReference("A1"));
        System.out.println(new CellReference("B"));
        System.out.println(new CellReference("2"));
    }
}
