package cc.whohow.excel;

import java.util.Objects;

public class ColumnKey {
    private String name;
    private String description;

    public ColumnKey() {
    }

    public ColumnKey(String name, String description) {
        this.name = name;
        this.description = description;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        ColumnKey columnKey = (ColumnKey) o;
        return Objects.equals(name, columnKey.name) &&
                Objects.equals(description, columnKey.description);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, description);
    }

    @Override
    public String toString() {
        return "ColumnKey{" +
                "name='" + name + '\'' +
                ", description='" + description + '\'' +
                '}';
    }
}
