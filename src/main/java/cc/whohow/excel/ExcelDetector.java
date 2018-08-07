package cc.whohow.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

public class ExcelDetector {
    private static final CellRangeAddress NULL = new CellRangeAddress(0, 0, 0, 0);
    private String headerSeparator = "\n";
    private Excel excel;
    private List<ColumnKey> keys;
    private CellRangeAddress headerRangeAddress;
    private CellRangeAddress bodyRangeAddress;

    public ExcelDetector(Excel excel) {
        this.excel = excel;
    }

    public boolean hasKeys() {
        return !(keys == null || keys.isEmpty());
    }

    public boolean hasHeader() {
        return headerRangeAddress != null && headerRangeAddress != NULL;
    }

    public List<ColumnKey> getKeys() {
        return keys;
    }

    public void setKeys(List<ColumnKey> keys) {
        this.keys = keys;
    }

    public CellRangeAddress getHeaderRangeAddress() {
        return headerRangeAddress;
    }

    public void setHeaderRangeAddress(CellRangeAddress headerRangeAddress) {
        this.headerRangeAddress = headerRangeAddress;
    }

    public void setBodyRangeAddress(CellRangeAddress bodyRangeAddress) {
        this.bodyRangeAddress = bodyRangeAddress;
    }

    public CellRangeAddress getBodyRangeAddress() {
        return bodyRangeAddress;
    }

    public void detectKeys() {
        if (hasKeys()) {
            return;
        }
        if (headerRangeAddress == null) {
            keys = new ExcelColumnKeys();
            return;
        }

        keys = new ArrayList<>();
        for (int c = headerRangeAddress.getFirstColumn(), i = 0; c <= headerRangeAddress.getLastColumn(); c++, i++) {
            Cell cell = excel.getCell(headerRangeAddress.getLastRow(), c);
            String text = excel.formatCellValue(cell, null);
            if (text == null || text.isEmpty()) {
                break;
            }
            keys.add(new ColumnKey(text, text, c));
        }
    }

    public void detectHeaderRangeAddress() {
        if (headerRangeAddress != null) {
            return;
        }
        CellRangeAddress sheetRangeAddress = excel.getSheetRangeAddress();
        for (int r = sheetRangeAddress.getFirstRow(); r <= sheetRangeAddress.getLastRow(); r++) {
            CellRangeAddress range = new CellRangeAddress(
                    r, r,
                    sheetRangeAddress.getFirstColumn(), sheetRangeAddress.getLastColumn());
            range = excel.trimRight(range);

            List<String> text = Arrays.asList(excel.getText(range, headerSeparator, null));
            Map<String, Integer> keysIndex = matchKeys(keys, text);
            if (keysIndex.size() == keys.size()) {
                headerRangeAddress = range;
                break;
            }
        }
    }

    public void detectBodyRangeAddress() {
        if (bodyRangeAddress != null) {
            return;
        }
        CellRangeAddress sheetRangeAddress = excel.getSheetRangeAddress();
        if (headerRangeAddress == null) {
            bodyRangeAddress = sheetRangeAddress;
        } else {
            bodyRangeAddress = new CellRangeAddress(
                    headerRangeAddress.getLastRow() + 1, sheetRangeAddress.getLastRow(),
                    headerRangeAddress.getFirstColumn(), headerRangeAddress.getLastColumn());
        }
    }

    public void detectKeysIndex() {
        if (headerRangeAddress == null) {
            int index = excel.getSheetRangeAddress().getFirstColumn();
            for (ColumnKey key : keys) {
                if (key.getIndex() >= 0) {
                    index = key.getIndex();
                } else {
                    index++;
                    key.setIndex(index);
                }
            }
        } else {
            List<String> text = Arrays.asList(excel.getText(headerRangeAddress, headerSeparator, null));
            Map<String, Integer> keysIndex = matchKeys(keys, text);
            for (ColumnKey key : keys) {
                if (key.getIndex() >= 0) {
                    continue;
                }
                Integer index = keysIndex.get(key.getName());
                if (index != null) {
                    key.setIndex(index);
                }
            }
        }
    }

    protected Map<String, Integer> matchKeys(List<ColumnKey> keys, List<String> text) {
        Map<String, Integer> keyIndex = new HashMap<>();
        for (ColumnKey key : keys) {
            int index = matchKey(key, text);
            if (index >= 0) {
                keyIndex.put(key.getName(), index);
            }
        }
        return keyIndex;
    }

    protected int matchKey(ColumnKey key, List<String> text) {
        int index = text.indexOf(key.getDescription());
        if (index < 0) {
            index = text.indexOf(key.getName());
        }
        return index;
    }
}
