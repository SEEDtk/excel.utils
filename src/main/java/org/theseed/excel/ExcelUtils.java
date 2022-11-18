/**
 *
 */
package org.theseed.excel;

import java.util.Iterator;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * This is a simple static class that contains useful Excel functions.
 *
 * @author Bruce Parrello
 *
 */
public class ExcelUtils {

    /**
     * @return the string in a cell, or an empty string if it has none
     *
     * @param cell	spreadsheet cell to examine
     */
    public static String stringValue(Cell cell) {
        String retVal = "";
        if (cell != null) {
            switch (cell.getCellType()) {
            case STRING :
            case FORMULA :
                retVal = StringUtils.trim(cell.getStringCellValue());
                break;
            case NUMERIC :
                retVal = Double.toString(cell.getNumericCellValue());
                break;
            default:
                break;
            }
        }
        return retVal;
    }

    /**
     * @return the number in a cell, or NaN if it has none
     *
     * @param cell	spreadsheet cell to examine
     */
    public static double numValue(Cell cell) {
        double retVal = Double.NaN;
        if (cell != null) {
            switch (cell.getCellType()) {
            case NUMERIC :
            case FORMULA :
                retVal = cell.getNumericCellValue();
                break;
            default:
                break;
            }
        }
        return retVal;
    }

    /**
     * @return the cell in the specified row and column
     *
     * @param sheet		worksheet containing the cell
     * @param row		row index (0-based)
     * @param col		column index (0-based)
     */
    public static Cell getCell(Sheet sheet, int row, int col) {
        Cell retVal = null;
        Row rowObject = sheet.getRow(row);
        if (rowObject != null)
            retVal = rowObject.getCell(col);
        return retVal;
    }

    /**
     * Cycle the row iterator until it is positioned after the specified marker row.  The marker
     * row is indicated by a specific value in the first column.
     *
     * @param rowIter	spreadsheet row iterator
     * @param marker	marker text
     *
     * @return TRUE if successful, FALSE if the marker was not found
     */
    public static boolean findMarker(Iterator<Row> rowIter, String marker) {
        boolean retVal = false;
        while (rowIter.hasNext() && ! retVal) {
            Cell cell = rowIter.next().getCell(0);
            String value = stringValue(cell);
            if (value.contentEquals(marker))
                retVal = true;
        }
        return retVal;
    }

}
