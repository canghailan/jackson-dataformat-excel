package cc.whohow.excel;

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

    protected boolean hasKeys() {
        return keys != null && !keys.isEmpty();
    }

    protected boolean hasHeader() {
        return headerRangeAddress != null && headerRangeAddress != NULL;
    }

    public List<ColumnKey> getKeys() {
        return hasKeys() ? keys : new ArrayList<>();
    }

    public void setKeys(List<ColumnKey> keys) {
        this.keys = keys;
    }

    public CellRangeAddress getHeaderRangeAddress() {
        return hasHeader() ? headerRangeAddress : null;
    }

    public void setHeaderRangeAddress(CellRangeAddress headerRangeAddress) {
        this.headerRangeAddress = headerRangeAddress;
    }

    public CellRangeAddress getBodyRangeAddress() {
        return bodyRangeAddress;
    }

    public void setBodyRangeAddress(CellRangeAddress bodyRangeAddress) {
        this.bodyRangeAddress = bodyRangeAddress;
    }

    public void detectKeys() {
        if (hasKeys()) {
            return;
        }
        if (headerRangeAddress == null) {
            keys = new ExcelColumnKeys();
            headerRangeAddress = NULL;
            return;
        }

        keys = new ArrayList<>();
        int c = headerRangeAddress.getFirstColumn();
        for (String text : getHeaderText(headerRangeAddress)) {
            if (text == null || text.isEmpty()) {
                break;
            }
            keys.add(new ColumnKey(text, text, c++));
        }
    }

    public void detectHeaderRangeAddress() {
        if (headerRangeAddress != null) {
            headerRangeAddress = excel.trimRight(headerRangeAddress);
            return;
        }
        CellRangeAddress sheetRangeAddress = excel.getSheetRangeAddress();
        for (int r = sheetRangeAddress.getFirstRow(); r <= sheetRangeAddress.getLastRow(); r++) {
            CellRangeAddress range = new CellRangeAddress(
                    r, r,
                    sheetRangeAddress.getFirstColumn(), sheetRangeAddress.getLastColumn());

            Map<String, Integer> keysIndex = matchKeys(keys, getHeaderText(range));
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
        if (hasHeader()) {
            bodyRangeAddress = new CellRangeAddress(
                    headerRangeAddress.getLastRow() + 1, sheetRangeAddress.getLastRow(),
                    headerRangeAddress.getFirstColumn(), headerRangeAddress.getLastColumn());
        } else {
            bodyRangeAddress = sheetRangeAddress;
        }
    }

    public void detectKeysIndex() {
        int index = excel.getSheetRangeAddress().getFirstColumn() - 1;
        if (hasHeader()) {
            Map<String, Integer> keysIndex = matchKeys(keys, getHeaderText(headerRangeAddress));
            for (ColumnKey key : keys) {
                if (key.getIndex() >= 0) {
                    index = key.getIndex();
                    continue;
                }
                Integer keyIndex = keysIndex.get(key.getName());
                if (keyIndex != null) {
                    index = keyIndex;
                    key.setIndex(index);
                } else {
                    index++;
                    key.setIndex(index);
                }
            }
        } else {
            for (ColumnKey key : keys) {
                if (key.getIndex() >= 0) {
                    index = key.getIndex();
                } else {
                    index++;
                    key.setIndex(index);
                }
            }
        }
    }

    protected Map<String, Integer> matchKeys(List<ColumnKey> keys, List<String> text) {
        Map<String, Integer> keyIndex = new HashMap<>();
        for (ColumnKey key : keys) {
            if (key.getIndex() >= 0) {
                keyIndex.put(key.getName(), key.getIndex());
                continue;
            }
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

    protected List<String> getHeaderText(CellRangeAddress range) {
        return Arrays.asList(excel.getText(excel.trimRight(range), headerSeparator, ""));
    }
}
