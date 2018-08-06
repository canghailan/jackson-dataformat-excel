package cc.whohow.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

public class ExcelDetector {
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
            List<String> text = Arrays.asList(excel.getRowText(excel.getRow(r), null));
            Map<String, Integer> keysIndex = matchKeys(keys, text);
            if (keysIndex.size() == keys.size()) {
                int firstColumn = keysIndex.values().stream()
                        .mapToInt(Integer::intValue)
                        .min()
                        .orElse(sheetRangeAddress.getFirstColumn());
                int lastColumn = keysIndex.values().stream()
                        .mapToInt(Integer::intValue)
                        .max()
                        .orElse(sheetRangeAddress.getLastColumn());
                headerRangeAddress = new CellRangeAddress(r, r, firstColumn, lastColumn);
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
            List<String> text = Arrays.asList(excel.getRowText(headerRangeAddress, null));
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
