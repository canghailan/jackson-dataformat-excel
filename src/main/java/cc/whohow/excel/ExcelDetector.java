package cc.whohow.excel;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;
import java.util.concurrent.Callable;

/**
 * Excel布局自动推测
 */
public class ExcelDetector implements Callable<Boolean> {
    private static final List<ColumnKey> AUTO_KEYS = Collections.emptyList();
    private static final CellRangeAddress AUTO_RANGE = new CellRangeAddress(0, 0, 0, 0);
    private Excel excel;
    private String headerSeparator;
    private List<ColumnKey> keys;
    private CellRangeAddress headerRangeAddress;
    private CellRangeAddress bodyRangeAddress;

    public ExcelDetector(Excel excel) {
        this.excel = excel;
        this.headerSeparator = "\r\n";
        this.keys = AUTO_KEYS;
        this.headerRangeAddress = AUTO_RANGE;
        this.bodyRangeAddress = AUTO_RANGE;
    }

    public ExcelDetector withKeys(List<ColumnKey> keys) {
        if (keys != null && !keys.isEmpty()) {
            this.keys = keys;
        }
        return this;
    }

    public ExcelDetector withHeaderRangeAddress(CellRangeAddress headerRangeAddress) {
        if (headerRangeAddress != null) {
            this.headerRangeAddress = headerRangeAddress;
        }
        return this;
    }

    public ExcelDetector withBodyRangeAddress(CellRangeAddress bodyRangeAddress) {
        if (bodyRangeAddress != null) {
            this.bodyRangeAddress = bodyRangeAddress;
        }
        return this;
    }

    public boolean isAutoKeys() {
        return keys == AUTO_KEYS;
    }

    public boolean isAutoHeader() {
        return headerRangeAddress == AUTO_RANGE;
    }

    public boolean isAutoBody() {
        return bodyRangeAddress == AUTO_RANGE;
    }

    public String getHeaderSeparator() {
        return headerSeparator;
    }

    public void setHeaderSeparator(String headerSeparator) {
        this.headerSeparator = headerSeparator;
    }

    public List<ColumnKey> getKeys() {
        return isAutoKeys() ? new ArrayList<>() : keys;
    }

    public void setKeys(List<ColumnKey> keys) {
        this.keys = keys;
    }

    public CellRangeAddress getHeaderRangeAddress() {
        return isAutoHeader() ? null : headerRangeAddress;
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

    @Override
    public Boolean call() {
        if (!isAutoKeys() && !isAutoHeader() && !isAutoBody()) {
            return true;
        }
        CellRangeAddress sheetRangeAddress = excel.getAndTrimSheetRangeAddress();
        if (isAutoHeader()) {
            // 如果Header需要推测
            if (!isAutoKeys()) {
                // 优先使用Keys进行推测
                detectHeaderByKeys(sheetRangeAddress);
            } else if (!isAutoBody()) {
                // 否则使用Body进行推测
                detectHeaderByBody(sheetRangeAddress);
            } else {
                // 默认Header
                usingDefaultHeader(sheetRangeAddress);
            }
        }
        if (isAutoHeader()) {
            // Header推测失败
            return false;
        }
        if (isAutoBody()) {
            // 如果Body需要推测，使用Header信息进行推测
            detectBodyByHeader(sheetRangeAddress);
        }
        if (isAutoKeys()) {
            // 如果Keys需要推测，使用Header信息进行推测
            detectKeysByHeader();
        } else if (isKeysNeedUpdate()) {
            // 更新Keys
            updateKeysByHeader();
        }
        return true;
    }

    protected void detectHeaderByKeys(CellRangeAddress sheetRangeAddress) {
        // 根据Keys推测Header：逐行扫描，最匹配Keys的行作为Header
        Map<String, Integer> keysIndex = Collections.emptyMap();
        for (int r = sheetRangeAddress.getFirstRow(); r <= sheetRangeAddress.getLastRow(); r++) {
            CellRangeAddress range = new CellRangeAddress(
                    r, r,
                    sheetRangeAddress.getFirstColumn(), sheetRangeAddress.getLastColumn());

            Map<String, Integer> matchKeysIndex = matchKeys(range);
            if (matchKeysIndex.size() > keysIndex.size()) {
                headerRangeAddress = range;
                keysIndex = matchKeysIndex;
                if (matchKeysIndex.size() == keys.size()) {
                    break;
                }
            }
        }
        if (acceptKeysIndex(keysIndex)) {
            // 采纳匹配结果，更新Keys
            for (ColumnKey key : keys) {
                if (key.getIndex() < 0) {
                    key.setIndex(keysIndex.getOrDefault(key.getName(), -1));
                }
            }
        } else {
            headerRangeAddress = AUTO_RANGE;
        }
    }

    protected boolean acceptKeysIndex(Map<String, Integer> keysIndex) {
        return !keysIndex.isEmpty();
    }

    protected Map<String, Integer> matchKeys(CellRangeAddress headerRangeAddress) {
        // 行匹配Keys
        List<String> headers = getHeaders(headerRangeAddress);
        Map<String, Integer> keyIndex = new HashMap<>();
        for (ColumnKey key : keys) {
            if (key.getIndex() >= 0) {
                keyIndex.put(key.getName(), key.getIndex());
                continue;
            }
            int index = matchKey(key, headers);
            if (index >= 0) {
                keyIndex.putIfAbsent(key.getName(), headerRangeAddress.getFirstColumn() + index);
            }
        }
        return keyIndex;
    }

    protected int matchKey(ColumnKey key, List<String> header) {
        // 列匹配Key
        int index = header.indexOf(key.getDescription());
        if (index < 0) {
            index = header.indexOf(key.getName());
        }
        return index;
    }

    protected void detectHeaderByBody(CellRangeAddress sheetRangeAddress) {
        // 根据Body推测Header，从Sheet第一行到Body的上一行为Header
        headerRangeAddress = new CellRangeAddress(
                sheetRangeAddress.getFirstRow(), bodyRangeAddress.getFirstRow() - 1,
                bodyRangeAddress.getFirstColumn(), bodyRangeAddress.getLastColumn());
    }

    protected void usingDefaultHeader(CellRangeAddress sheetRangeAddress) {
        // 默认Header：Sheet第一行为默认Header
        headerRangeAddress = new CellRangeAddress(
                sheetRangeAddress.getFirstRow(), sheetRangeAddress.getFirstRow(),
                sheetRangeAddress.getFirstColumn(), sheetRangeAddress.getLastColumn());
    }

    protected void detectBodyByHeader(CellRangeAddress sheetRangeAddress) {
        // 根据Header推测Body：Header的下一行到Sheet最后一行为Body
        bodyRangeAddress = new CellRangeAddress(
                headerRangeAddress.getLastRow() + 1, sheetRangeAddress.getLastRow(),
                headerRangeAddress.getFirstColumn(), headerRangeAddress.getLastColumn());
    }

    protected void detectKeysByHeader() {
        // 根据Header推测Keys：Header的文本为Key的name和description
        List<String> headers = getHeaders(headerRangeAddress);
        keys = new ArrayList<>(headers.size());
        for (int i = 0; i < headers.size(); i++) {
            String header = headers.get(i);
            keys.add(new ColumnKey(header, header, headerRangeAddress.getFirstColumn() + i));
        }
    }

    protected boolean isKeysNeedUpdate() {
        // Keys是否需要更新：如果存在index未确定，则需要更新
        return keys.stream()
                .anyMatch(key -> key.getIndex() < 0);
    }

    protected void updateKeysByHeader() {
        // 根据Header更新Keys索引
        List<String> headers = getHeaders(headerRangeAddress);
        for (ColumnKey key : keys) {
            if (key.getIndex() < 0) {
                key.setIndex(matchKey(key, headers));
            }
        }
    }

    protected List<String> getHeaders(CellRangeAddress headerRangeAddress) {
        // 读取Header的文本
        return Arrays.asList(excel.getText(headerRangeAddress, headerSeparator, ""));
    }
}
