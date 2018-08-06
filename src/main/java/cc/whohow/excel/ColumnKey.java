package cc.whohow.excel;

import java.util.Objects;

public class ColumnKey {
    private final String name;
    private String description;
    private int index;

    public ColumnKey(String name) {
        this(name, name, -1);
    }

    public ColumnKey(String name, String description) {
        this(name, description, -1);
    }

    public ColumnKey(String name, String description, int index) {
        this.name = name;
        this.description = description;
        this.index = index;
    }

    public String getName() {
        return name;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        ColumnKey columnKey = (ColumnKey) o;
        return Objects.equals(name, columnKey.name);
    }

    @Override
    public int hashCode() {
        return Objects.hashCode(name);
    }

    @Override
    public String toString() {
        return name;
    }
}
