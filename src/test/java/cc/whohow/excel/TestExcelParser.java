package cc.whohow.excel;

import com.fasterxml.jackson.core.JsonToken;
import org.junit.Test;

import java.io.File;

public class TestExcelParser {
    @Test
    public void test() throws Exception {
        ExcelFactory factory = new ExcelFactory();
        try (ExcelParser parser = factory.createParser(new File("test.xlsx"))) {
//            ExcelSchema schema = new ExcelSchema()
//                    .withHeader(null);
//            parser.setSchema(schema);
            while (true) {
                JsonToken token = parser.nextToken();
                if (token == null) {
                    break;
                }
                System.out.println(parser.getText());
            }
        }
    }
}
