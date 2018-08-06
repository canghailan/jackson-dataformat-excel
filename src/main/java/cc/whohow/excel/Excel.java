package cc.whohow.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel {
    private static final Pattern CELL_RANGE_ADDRESS = Pattern
            .compile("(?<firstColumn>[A-Z]*)(?<firstRow>[0-9]*):?(?<lastColumn>[A-Z]*)(?<lastRow>[0-9]*)");

    protected final ISO8601VariantDateFormat dateFormat = new ISO8601VariantDateFormat();
    protected final Workbook workbook;
    protected final Sheet sheet;
    protected final FormulaEvaluator formulaEvaluator;
    protected NumberFormat numberFormat;

    public Excel() {
        this(SpreadsheetVersion.EXCEL2007);
    }

    public Excel(Sheet sheet) {
        this.workbook = sheet.getWorkbook();
        this.sheet = sheet;
        this.formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        this.numberFormat = new DecimalFormat("#.#");
        this.numberFormat.setGroupingUsed(false);
    }

    public Excel(Workbook workbook) {
        this(workbook.getSheetAt(workbook.getActiveSheetIndex()));
    }

    public Excel(SpreadsheetVersion version) {
        this(version == SpreadsheetVersion.EXCEL97 ? new HSSFWorkbook() : new XSSFWorkbook());
    }

    public CellRangeAddress getSheetRangeAddress() {
        return new CellRangeAddress(
                sheet.getFirstRowNum(), sheet.getLastRowNum(),
                sheet.getLeftCol(), workbook.getSpreadsheetVersion().getLastColumnIndex());
    }

    public CellRangeAddress getTrimmedSheetRangeAddress() {
        return null;
    }

    public CellRangeAddress getRowRangeAddress(Row row) {
        return null;
    }

    public CellRangeAddress getTrimmedRowRangeAddress(Row row) {
        return null;
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

    public CellRangeAddress getTrimmedCellRangeAddress(String ref) {
        return getCellRangeAddress(ref);
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
        return normalizeCellValue(getCellWithMerges(row.getRowNum(), column));
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

    public String[] getRowText(Row row, String defaultValue) {
        if (row == null || row.getLastCellNum() <= 0) {
            return new String[0];
        }
        String[] result =  new String[row.getLastCellNum()];
        for (int c = 0; c < row.getLastCellNum(); c++) {
            result[c] = formatCellValue(getCell(row, c), defaultValue);
        }
        return result;
    }

    public String[] getRowText(CellRangeAddress range, String defaultValue) {
        return null;
    }

    public String[] getRowText(Row row, CellRangeAddress range, String defaultValue) {
        CellRangeAddress trimmedRowRangeAddress = getTrimmedRowRangeAddress(row);
        int first = range.getFirstColumn();
        int last = range.getLastColumn();
        if (last > trimmedRowRangeAddress.getLastColumn()) {
            last = trimmedRowRangeAddress.getLastColumn();
        }
        if (last < first) {
            return new String[0];
        }

        String[] result = new String[last - first + 1];
        for (int i = 0, c = first; i < result.length; i++, c++) {
            result[i] = formatCellValue(getCell(row, c), defaultValue);
        }
        return result;
    }

    protected Object getCellValue(Cell cell) {
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

    protected String formatCellValue(Cell cell) {
        return formatCellValue(cell, null);
    }

    protected String formatCellValue(Cell cell, String defaultValue) {
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

    protected String format(Date dateValue) {
        return dateFormat.format(dateValue);
    }

    protected String format(double numericValue) {
        return numberFormat.format(numericValue);
    }

    protected String format(boolean booleanValue) {
        return Boolean.toString(booleanValue);
    }

    protected boolean isEmptyCell(Cell cell) {
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

    protected boolean isEmptyRow(Row row, CellRangeAddress range) {
        if (row == null) {
            return true;
        }
        int first = range.getFirstColumn();
        if (first < row.getFirstCellNum()) {
            first = row.getFirstCellNum();
        }
        int last = range.getLastColumn();
        if (last >= row.getLastCellNum()) {
            last = row.getLastCellNum() - 1;
        }
        for (int c = first; c <= last; c++) {
            if (!isEmptyCell(getCell(row.getRowNum(), c))) {
                return false;
            }
        }
        return true;
    }
}
