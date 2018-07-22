package cc.whohow.excel;

import java.util.Objects;

public class ColumnKey {
    private final String name;
    private String description;

    public ColumnKey(String name) {
        this(name, null);
    }

    public ColumnKey(String name, String description) {
        this.name = name;
        this.description = description;
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
        if (description == null || description.isEmpty()) {
            return name;
        }
        return name + "(" + description + ")";
    }
}
