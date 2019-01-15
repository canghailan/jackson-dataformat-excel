package cc.whohow.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel {
    private static final Pattern CELL_RANGE_ADDRESS = Pattern
            .compile("(?<firstColumn>[A-Z]*)(?<firstRow>[0-9]*):?(?<lastColumn>[A-Z]*)(?<lastRow>[0-9]*)");

    protected final Workbook workbook;
    protected final Sheet sheet;
    protected final FormulaEvaluator formulaEvaluator;
    protected DateFormat dateFormat;
    protected NumberFormat numberFormat;

    public Excel() {
        this(SpreadsheetVersion.EXCEL2007);
    }

    public Excel(Sheet sheet) {
        this.workbook = sheet.getWorkbook();
        this.sheet = sheet;
        this.formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        this.dateFormat = new ISO8601VariantDateFormat();
        this.numberFormat = new DecimalFormat("#.#");
        this.numberFormat.setGroupingUsed(false);
    }

    public Excel(Workbook workbook) {
        this(workbook.getActiveSheetIndex() < workbook.getNumberOfSheets() ?
                workbook.getSheetAt(workbook.getActiveSheetIndex()) :
                workbook.createSheet());
    }

    public Excel(SpreadsheetVersion version) {
        this(version == SpreadsheetVersion.EXCEL97 ? new HSSFWorkbook() : new XSSFWorkbook());
    }

    public DateFormat getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(DateFormat dateFormat) {
        this.dateFormat = dateFormat;
    }

    public NumberFormat getNumberFormat() {
        return numberFormat;
    }

    public void setNumberFormat(NumberFormat numberFormat) {
        this.numberFormat = numberFormat;
    }

    public CellRangeAddress getSheetRangeAddress() {
        return new CellRangeAddress(
                sheet.getFirstRowNum(), sheet.getLastRowNum(),
                sheet.getLeftCol(), workbook.getSpreadsheetVersion().getLastColumnIndex());
    }

    public CellRangeAddress getAndTrimSheetRangeAddress() {
        return trim(getSheetRangeAddress());
    }

    public CellRangeAddress getRowRangeAddress(int row) {
        return getRowRangeAddress(getRow(row));
    }

    public CellRangeAddress getRowRangeAddress(Row row) {
        if (row == null) {
            return null;
        }
        if (row.getFirstCellNum() < 0) {
            return new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 0);
        }
        return new CellRangeAddress(row.getRowNum(), row.getRowNum(), row.getFirstCellNum(), row.getLastCellNum() - 1);
    }

    public CellRangeAddress getCellRangeAddress(String ref) {
        if (ref == null) {
            return null;
        }
        Matcher matcher = CELL_RANGE_ADDRESS.matcher(ref);
        if (matcher.matches()) {
            int firstRowIndex = sheet.getFirstRowNum();
            int firstColumnIndex = sheet.getLeftCol();
            int lastRowIndex = sheet.getLastRowNum();
            int lastColumnIndex = workbook.getSpreadsheetVersion().getLastColumnIndex();

            String firstColumn = matcher.group(1);
            String firstRow = matcher.group(2);
            String lastColumn = matcher.group(3);
            String lastRow = matcher.group(4);

            if (!firstRow.isEmpty()) {
                firstRowIndex = Integer.parseInt(firstRow) - 1;
            }
            if (!firstColumn.isEmpty()) {
                firstColumnIndex = CellReference.convertColStringToIndex(firstColumn);
            }
            if (!lastRow.isEmpty()) {
                lastRowIndex = Integer.parseInt(lastRow) - 1;
            }
            if (!lastColumn.isEmpty()) {
                lastColumnIndex = CellReference.convertColStringToIndex(lastColumn);
            }

            if (lastRowIndex < firstRowIndex) {
                lastRowIndex = firstRowIndex;
            }
            if (lastColumnIndex < firstColumnIndex) {
                lastColumnIndex = firstColumnIndex;
            }
            return new CellRangeAddress(
                    firstRowIndex, lastRowIndex,
                    firstColumnIndex, lastColumnIndex);
        }
        throw new IllegalArgumentException("CellRangeAddress: " + ref);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Row getRow(int row) {
        return sheet.getRow(row);
    }

    public Row createRow(int row) {
        return CellUtil.getRow(row, sheet);
    }

    public Cell getCell(Row row, int column) {
        return row == null ? null : getCell(row.getRowNum(), column);
    }

    public Cell getCell(int row, int column) {
        return normalizeCellValue(getCellWithMerges(row, column));
    }

    public Cell getCellWithMerges(int row, int column) {
        return SheetUtil.getCellWithMerges(sheet, row, column);
    }

    public Cell normalizeCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        if (cell.getCellTypeEnum() == CellType.FORMULA) {
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            cell.setCellValue(cellValue.getNumberValue());
            cell.setCellValue(cellValue.getBooleanValue());
            cell.setCellValue(cellValue.getStringValue());
            cell.setCellErrorValue(cellValue.getErrorValue());
            cell.setCellType(cellValue.getCellTypeEnum());
        }
        if (cell.getCellTypeEnum() == CellType.NUMERIC &&
                DateUtil.isCellDateFormatted(cell)) {
            cell.setCellValue(format(cell.getDateCellValue()));
            cell.setCellType(CellType.STRING);
        }
        return cell;
    }

    public Cell createCell(int row, int column) {
        return createCell(createRow(row), column);
    }

    public Cell createCell(Row row, int column) {
        return CellUtil.getCell(row, column);
    }

    public String getText(int row, int column) {
        return formatCellValue(getCell(row, column));
    }

    public String[][] getText(CellRangeAddress range, String defaultValue) {
        int rows = range.getLastRow() - range.getFirstRow() + 1;
        int columns = range.getLastColumn() - range.getFirstColumn() + 1;
        String[][] result = new String[rows][columns];
        for (int i = 0, r = range.getFirstRow(); i < result.length; i++, r++) {
            for (int j = 0, c = range.getFirstColumn(); j < result[i].length; j++, c++) {
                result[i][j] = formatCellValue(getCell(r, c), defaultValue);
            }
        }
        return result;
    }

    public String[] getText(CellRangeAddress range, String separator, String defaultValue) {
        String[] result = new String[range.getLastColumn() - range.getFirstColumn() + 1];
        if (range.getFirstRow() == range.getLastRow()) {
            for (int i = 0, c = range.getFirstColumn(); i < result.length; i++, c++) {
                result[i] = formatCellValue(getCell(range.getFirstRow(), c), defaultValue);
            }
            return result;
        } else {
            for (int i = 0, c = range.getFirstColumn(); i < result.length; i++, c++) {
                String[] rows = new String[range.getLastRow() - range.getFirstRow() + 1];
                for (int j = 0, r = range.getFirstRow(); j < rows.length; j++, r++) {
                    rows[j] = formatCellValue(getCell(r, c), defaultValue);
                }
                result[i] = String.join(separator, rows);
            }
            return result;
        }
    }

    public String[] getText(Row row, String defaultValue) {
        if (row == null || row.getLastCellNum() <= 0) {
            return new String[0];
        }
        String[] result = new String[row.getLastCellNum()];
        for (int c = 0; c < row.getLastCellNum(); c++) {
            result[c] = formatCellValue(getCell(row, c), defaultValue);
        }
        return result;
    }

    public Object getValue(int row, int column) {
        return getCellValue(getCell(row, column));
    }

    public Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellTypeEnum()) {
            case STRING: {
                return cell.getStringCellValue();
            }
            case NUMERIC: {
                return cell.getNumericCellValue();
            }
            case BOOLEAN: {
                return cell.getBooleanCellValue();
            }
            default: {
                return null;
            }
        }
    }

    public String formatCellValue(Cell cell) {
        return formatCellValue(cell, null);
    }

    public String formatCellValue(Cell cell, String defaultValue) {
        if (cell == null) {
            return defaultValue;
        }
        switch (cell.getCellTypeEnum()) {
            case STRING: {
                return cell.getStringCellValue();
            }
            case NUMERIC: {
                return format(cell.getNumericCellValue());
            }
            case BOOLEAN: {
                return format(cell.getBooleanCellValue());
            }
            default: {
                return defaultValue;
            }
        }
    }

    public String format(Date dateValue) {
        return (dateValue == null) ? null : dateFormat.format(dateValue);
    }

    public String format(double numericValue) {
        return numberFormat.format(numericValue);
    }

    public String format(boolean booleanValue) {
        return Boolean.toString(booleanValue);
    }

    public boolean isEmptyCell(Cell cell) {
        if (cell == null) {
            return true;
        }
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
            case BOOLEAN: {
                return false;
            }
            case STRING: {
                return cell.getStringCellValue().isEmpty();
            }
            default: {
                return true;
            }
        }
    }

    public boolean isEmptyRow(Row row) {
        return isEmpty(getRowRangeAddress(row));
    }

    public boolean isEmpty(CellRangeAddress range) {
        if (range == null) {
            return true;
        }
        for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
            Row row = getRow(r);
            if (row == null || row.getLastCellNum() < 0) {
                continue;
            }
            int firstColumn = Integer.max(range.getFirstColumn(), row.getFirstCellNum());
            int lastColumn = Integer.min(range.getLastColumn(), row.getLastCellNum() - 1);
            for (int c = firstColumn; c <= lastColumn; c++) {
                if (!isEmptyCell(getCell(row, c))) {
                    return false;
                }
            }
        }
        return true;
    }

    public CellRangeAddress trim(CellRangeAddress range) {
        return trimLeft(trimRight(trimTop(trimBottom(range))));
    }

    public CellRangeAddress trimTop(CellRangeAddress range) {
        if (range == null) {
            return null;
        }
        int firstRow = range.getFirstRow();
        while (firstRow < range.getLastRow()) {
            if (isEmpty(new CellRangeAddress(firstRow, firstRow, range.getFirstColumn(), range.getLastColumn()))) {
                firstRow++;
            } else {
                break;
            }
        }
        return new CellRangeAddress(firstRow, range.getLastRow(), range.getFirstColumn(), range.getLastColumn());
    }

    public CellRangeAddress trimBottom(CellRangeAddress range) {
        if (range == null) {
            return null;
        }
        int lastRow = range.getLastRow();
        while (lastRow > range.getFirstRow()) {
            if (isEmpty(new CellRangeAddress(lastRow, lastRow, range.getFirstColumn(), range.getLastColumn()))) {
                lastRow--;
            } else {
                break;
            }
        }
        return new CellRangeAddress(range.getFirstRow(), lastRow, range.getFirstColumn(), range.getLastColumn());
    }

    public CellRangeAddress trimLeft(CellRangeAddress range) {
        if (range == null) {
            return null;
        }
        int firstColumn = range.getFirstColumn();
        while (firstColumn < range.getLastColumn()) {
            if (isEmpty(new CellRangeAddress(range.getFirstRow(), range.getLastRow(), firstColumn, firstColumn))) {
                firstColumn++;
            } else {
                break;
            }
        }
        return new CellRangeAddress(range.getFirstRow(), range.getLastRow(), firstColumn, range.getLastColumn());
    }

    public CellRangeAddress trimRight(CellRangeAddress range) {
        if (range == null) {
            return null;
        }
        int lastColumn = range.getLastColumn();
        while (lastColumn > range.getFirstColumn()) {
            if (isEmpty(new CellRangeAddress(range.getFirstRow(), range.getLastRow(), lastColumn, lastColumn))) {
                lastColumn--;
            } else {
                break;
            }
        }
        return new CellRangeAddress(range.getFirstRow(), range.getLastRow(), range.getFirstColumn(), lastColumn);
    }

    /**
     * 求交集，无交集返回null
     */
    public CellRangeAddress intersect(CellRangeAddress range1, CellRangeAddress range2) {
        int firstRow = Integer.max(range1.getFirstRow(), range2.getFirstRow());
        int lastRow = Integer.min(range1.getLastRow(), range2.getLastRow());
        if (firstRow > lastRow) {
            return null;
        }
        int firstColumn = Integer.max(range1.getFirstColumn(), range2.getFirstColumn());
        int lastColumn = Integer.min(range1.getLastColumn(), range2.getLastColumn());
        if (firstColumn > lastColumn) {
            return null;
        }
        return new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
    }
}
