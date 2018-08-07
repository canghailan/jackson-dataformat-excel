package cc.whohow.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellReference;

import java.util.AbstractList;
import java.util.Arrays;
import java.util.Set;
import java.util.Spliterator;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class ExcelColumnKeys extends AbstractList<ColumnKey> implements Set<ColumnKey> {
    private final SpreadsheetVersion spreadsheetVersion;
    private ColumnKey[] cache = new ColumnKey[26];

    public ExcelColumnKeys() {
        this(SpreadsheetVersion.EXCEL2007);
    }

    public ExcelColumnKeys(SpreadsheetVersion spreadsheetVersion) {
        this.spreadsheetVersion = spreadsheetVersion;
    }

    @Override
    public ColumnKey get(int index) {
        ensureCacheCapacity(index + 1);
        if (cache[index] == null) {
            cache[index] = newColumnKey(index);
        }
        return cache[index];
    }

    @Override
    public Stream<ColumnKey> stream() {
        return IntStream.range(0, size())
                .mapToObj(this::newColumnKey);
    }

    @Override
    public Spliterator<ColumnKey> spliterator() {
        return stream().spliterator();
    }

    @Override
    public int size() {
        return spreadsheetVersion.getMaxColumns();
    }

    protected void ensureCacheCapacity(int minLength) {
        if (minLength <= cache.length) {
            return;
        }
        int length = cache.length;
        while (length < minLength) {
            length *= 2;
        }
        cache = Arrays.copyOf(cache, length);
    }

    protected ColumnKey newColumnKey(int index) {
        String columnRef = CellReference.convertNumToColString(index);
        return new ColumnKey(columnRef.toLowerCase(), columnRef, index);
    }
}
