/**
 *
 */
package org.theseed.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

/**
 * This object manages a workbook with tables and optionally various useful formattng options.   It supports
 * multiple sheets being filled at once, but also a single default sheet.  It is assumed
 * that the data rows will be added sequentially and the cells will be added left to right.  This makes
 * it easy to stream data into the sheets.
 *
 * @author Bruce Parrello
 *
 */
public class CustomWorkbook implements AutoCloseable {

    /**
     * This enum defines the style types for floating-point numbers.
     */
    public static enum Num {
        NORMAL, FRACTION, ML;
    }

    /**
     * This enum defines the style types for text.
     */
    public static enum Text {
        NORMAL, FLAG;
    }

    // FIELDS
    /** master workbook */
    private XSSFWorkbook workbook;
    /** default floating-point precision ("num") format */
    private int precision;
    /** maximum column width for the current sheet */
    private int maxWidth;
    /** file in which to save workbook */
    private File outFile;
    /** default worksheet for sheet-level operations */
    private Sheet defaultSheet;
    /** creation helper for links */
    private XSSFCreationHelper helper;
    /** map of table names */
    private TableNameMap tableMap;
    /** queue of sheets to delete */
    private List<XSSFSheet> deleteQueue;
    /** number of dead tables */
    private int deadTables;
    /** list of open sheets */
    private List<Sheet> openSheets;
    /** cell reference for spreadsheet origin */
    private static final CellReference ORIGIN_REF = new CellReference(0, 0);

    // STYLES
    /** normal number format */
    private XSSFCellStyle numStyle;
    /** fraction number format */
    private XSSFCellStyle fracStyle;
    /** machine-learning input number format */
    private XSSFCellStyle mlStyle;
    /** integer number format */
    private XSSFCellStyle intStyle;
    /** hyperlink cell format */
    private XSSFCellStyle linkStyle;
    /** high-number format */
    private XSSFCellStyle highStyle;
    /** low-number format */
    private XSSFCellStyle lowStyle;
    /** text format */
    private XSSFCellStyle textStyle;
    /** flag format */
    private XSSFCellStyle flagStyle;
    /** autowrap format */
    private XSSFCellStyle wrapStyle;
    /** linked autowrap format */
    private XSSFCellStyle lwrapStyle;
    /** header format */
    private XSSFCellStyle headStyle;
    /** workbook format helper */
    private DataFormat formatter;


    /**
     * This class represents a single worksheet.
     */
    public class Sheet implements AutoCloseable {

        // FIELDS
        /** current sheet */
        public XSSFSheet sheet;
        /** index of next row to add */
        public int rowIdx;
        /** current sheet header row */
        public XSSFRow headerRow;
        /** current sheet row */
        public XSSFRow row;
        /** index of next cell to add to the row */
        public int colIdx;
        /** maximum width of spreadsheet in cells */
        public int maxCols;
        /** TRUE if we want this sheet to be a table */
        public boolean tableMode;
        /** this will be an array of the required header widths */
        public int[] autoWidths;
        /** drawing helper for comments and graphs */
        private final XSSFDrawing drawHelper;

        /**
         * Create a new worksheet.
         *
         * @param name		name of the new worksheet
         * @param isTable	TRUE to make this worksheet a table
         */
        protected Sheet(String name, boolean isTable) {
            // Does this sheet already exist?
            XSSFSheet oldSheet = CustomWorkbook.this.workbook.getSheet(name);
            if (oldSheet != null) {
                // Yes.  Create a new sheet with a dummy name.
                this.sheet = CustomWorkbook.this.workbook.createSheet();
                // We must set up the old sheet for deletion.  We have to delete it at the
                // end because of a bug in apache.
                CustomWorkbook.this.deleteQueue.add(oldSheet);
                // Rename the tables to avoid conflicts.
                for (XSSFTable table : this.sheet.getTables()) {
                    CTTable cttable = table.getCTTable();
                    CustomWorkbook.this.deadTables++;
                    String deadName = String.format("__dead_table_%d", CustomWorkbook.this.deadTables);
                    cttable.setName(deadName);
                    cttable.setDisplayName(deadName);
                }
                // Give the old sheet a dead name.
                int oldIdx = CustomWorkbook.this.workbook.getSheetIndex(oldSheet);
                CustomWorkbook.this.workbook.setSheetName(oldIdx, String.format("_deadSheet%d", CustomWorkbook.this.deleteQueue.size()));
                // Rename the new sheet.  This is now the real one.
                int newIdx = CustomWorkbook.this.workbook.getSheetIndex(this.sheet);
                CustomWorkbook.this.workbook.setSheetName(newIdx, name);
            } else {
                // Sheet is new.  Create it the simple way.
                this.sheet = CustomWorkbook.this.workbook.createSheet(name);
            }
            // Create the header row.
            this.headerRow = this.sheet.createRow(0);
            // Position ourselves at the top of the sheet.
            this.drawHelper = this.sheet.createDrawingPatriarch();
            this.rowIdx = 1;
            this.maxCols = 0;
            this.tableMode = isTable;
        }

        /**
         * Store the headers for this worksheet.
         *
         * @param headers	list of header names to store
         */
        public void setHeaders(List<String> headers) {
            // Select the style.
            XSSFCellStyle style = (this.tableMode ? CustomWorkbook.this.textStyle : CustomWorkbook.this.headStyle);
            final int n = headers.size();
            this.autoWidths = new int[n];
            for (int i = 0; i < n; i++) {
                XSSFCell curr = this.headerRow.createCell(i, CellType.STRING);
                curr.setCellValue(headers.get(i));
                curr.setCellStyle(style);
                this.sheet.autoSizeColumn(i);
                this.autoWidths[i] = this.sheet.getColumnWidth(i) + (this.tableMode ? 512 : 16);
            }
        }

        /**
         * Store a range-colored value in the next cell of this row.  The value will be normally-colored
         * if it is between the minimum and maximum.  If it is at or below the minimum, it will be red.  If
         * it is at or above the maximum, it will be green.
         *
         * @param value		value to store
         * @param min		maximum "low" value
         * @param max		minimum "high" value
         */
        public void storeCell(double value, double min, double max) {
            XSSFCell cell = this.addCell();
            cell.setCellValue(value);
            if (value <= min)
                cell.setCellStyle(CustomWorkbook.this.lowStyle);
            else if (value >= max)
                cell.setCellStyle(CustomWorkbook.this.highStyle);
            else
                cell.setCellStyle(CustomWorkbook.this.numStyle);
        }

        /**
         * Add a new data row to the sheet.
         */
        public void addRow() {
            this.row = this.sheet.createRow(this.rowIdx);
            this.colIdx = 0;
            this.rowIdx++;
        }

        /**
         * Add a new cell to the sheet at the current position.
         *
         * @return the new cell added
         */
        private XSSFCell addCell() {
            XSSFCell retVal = this.row.createCell(this.colIdx);
            this.colIdx++;
            if (this.colIdx > this.maxCols) this.maxCols = this.colIdx;
            return retVal;
        }

        /**
         * Store a floating-point value in the next cell of this row.
         *
         * @param value		value to store
         * @param style		style of number
         */
        public void storeCell(double value, Num style) {
            XSSFCell cell = this.addCell();
            this.storeDouble(cell, value, style);
        }

        /**
         * Store a floating-point value in a cell.
         *
         * @param cell		target cell
         * @param value		value to store
         * @param style		style of number
         */
        private void storeDouble(XSSFCell cell, double value, Num style) {
            cell.setCellValue(value);
            switch (style) {
            case NORMAL -> cell.setCellStyle(CustomWorkbook.this.numStyle);
            case FRACTION -> cell.setCellStyle(CustomWorkbook.this.fracStyle);
            case ML -> cell.setCellStyle(CustomWorkbook.this.mlStyle);
            }
        }

        /**
         * Store a floating-point value in a random cell.  Random-access stores should be done after
         * the sequential storing that is the norm.
         *
         * @param r			target row index
         * @param c			target column index
         * @param value		value to store
         * @param style		style of number
         *
         */
        public void storeCell(int r, int c, double value, Num style) {
            XSSFCell cell = this.findCell(r, c);
            this.storeDouble(cell, value, style);
        }

        /**
         * @return the identified spreadsheet cell
         *
         * @param r		target row index
         * @param c		target column index
         */
        private XSSFCell findCell(int r, int c) {
            var myRow = this.sheet.getRow(r);
            if (myRow == null)
                myRow = this.sheet.createRow(r);
            return myRow.createCell(c);
        }

        /**
         * Store a string value in a random cell.  Random stores should be done after the sequential storing
         * that is the norm.
         *
         * @param r			target row index
         * @param c			target column index
         * @param value		value to store
         */
        public void storeCell(int r, int c, String value) {
            var cell = this.findCell(r, c);
            cell.setCellValue(value);
        }

        /**
         * Store an integer value in the next cell of this row.
         *
         * @param value		value to store
         */
        public void storeCell(int value) {
            XSSFCell cell = this.addCell();
            cell.setCellValue((double) value);
            cell.setCellStyle(CustomWorkbook.this.intStyle);
        }

        /**
         * Store a text value in the next cell of this row.
         *
         * @param value		value to store
         * @param style		style of text
         */
        public void storeCell(String value, Text style) {
            if (StringUtils.isBlank(value))
                this.addCell();
            else {
                // Here we have real text to store in the cell.
                XSSFCell cell = this.addCell();
                cell.setCellValue(value);
                switch (style) {
                case NORMAL -> cell.setCellStyle(CustomWorkbook.this.textStyle);
                case FLAG -> cell.setCellStyle(CustomWorkbook.this.flagStyle);
                }
            }
        }

        /**
         * Store a text value and link in the next cell of this row.
         *
         * @param value		value to store
         * @param url		URL for the link (NULL for none)
         * @param comment	comment for the cell (NULL for none)
         */
        public void storeCell(String value, String url, String comment) {
            if (StringUtils.isBlank(value))
                this.addCell();
            else {
                // Here we have real text to put in the cell.
                XSSFCell cell = this.addCell();
                cell.setCellValue(value);
                // Process the link and comment.
                this.decorate(cell, url, comment);
            }
        }

        /**
         * Decorate a cell with an optional link and comment.
         *
         * @param cell		cell to decorate
         * @param url		URL for the link, or NULL for no link
         * @param comment	text of the comment, or NULL for no comment
         */
        private void decorate(XSSFCell cell, String url, String comment) {
            if (StringUtils.isBlank(url)) {
                // No link was provided, so format the cell as text.
                cell.setCellStyle(CustomWorkbook.this.textStyle);
            } else {
                // Here we have the URL for a link.
                XSSFHyperlink link = CustomWorkbook.this.helper.createHyperlink(HyperlinkType.URL);
                link.setAddress(url);
                cell.setHyperlink(link);
                cell.setCellStyle(CustomWorkbook.this.linkStyle);
            }
            if (! StringUtils.isBlank(comment)) {
                // Here we have to add a comment.
                int r = cell.getRowIndex();
                int c = cell.getColumnIndex();
                // This describes where the comment appears.  It appears under the cell.  The first four 0s are
                // within-cell displacements.  We cover 5 columns and 2 rows.
                XSSFClientAnchor anchor = this.drawHelper.createAnchor(0, 0, 0, 0, c, r+1, c+5, r+3);
                XSSFComment commentObject = this.drawHelper.createCellComment(anchor);
                commentObject.setAddress(r, c);
                commentObject.setString(comment);
                cell.setCellComment(commentObject);
            }
        }

        /**
         * Store an empty cell in the current position.
         */
        public void storeBlankCell() {
            this.addCell();
        }

        /**
         * Store an integer cell with a hyperlink.
         *
         * @param value		integer value in the cell
         * @param url		URL for the link (or NULL if none)
         * @param comment	comment text (or NULL if no comment)
         */
        public void storeCell(int value, String url, String comment) {
            XSSFCell cell = this.addCell();
            cell.setCellValue((double) value);
            this.decorate(cell, url, comment);
        }

        /**
         * Store a string in a cell and format it normally.
         *
         * @param value		string to store
         */
        public void storeCell(String string) {
            this.storeCell(string, Text.NORMAL);
        }

        /**
         * Store a number in a cell and format it normally.
         *
         * @param value		number to store
         */
        public void storeCell(double value) {
            this.storeCell(value, Num.NORMAL);
        }

        /**
         * Reformat a number column as integer.
         *
         * @param c		index of the column to reformat
         */
        public void reformatIntColumn(int c) {
            for (int r = 1; r < this.rowIdx; r++) {
                XSSFCell cell = this.sheet.getRow(r).getCell(c);
                if (cell.getCellType() == CellType.NUMERIC)
                    cell.setCellStyle(CustomWorkbook.this.intStyle);
            }
        }

        /**
         * Autosize the specified column.
         *
         * @param c		index of the column to autosize
         */
        public void autoSizeColumn(int c) {
            this.sheet.autoSizeColumn(c);
            if (this.tableMode) {
                // Here we need to add space for the filter arrow.
                int cWidth = this.sheet.getColumnWidth(c);
                if (cWidth < this.autoWidths[c])
                    this.sheet.setColumnWidth(c, this.autoWidths[c]);
                else if (cWidth > CustomWorkbook.this.maxWidth) {
                    this.sheet.setColumnWidth(c, CustomWorkbook.this.maxWidth);
                    for (int r = 0; r < this.rowIdx; r++) {
                        XSSFCell cell = this.sheet.getRow(r).getCell(c);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            // We need to set the cell to wrap.  Does it have a link?
                            if (cell.getHyperlink() != null)
                                cell.setCellStyle(CustomWorkbook.this.lwrapStyle);
                            else
                                cell.setCellStyle(CustomWorkbook.this.wrapStyle);
                        }
                    }
                }
            }
        }

        /**
         * Reformat a text column as flags.
         *
         * @param c		index of column to reformat
         */
        public void reformatFlagColumn(int c) {
            for (int r = 1; r < this.rowIdx; r++) {
                XSSFCell cell = this.sheet.getRow(r).getCell(c);
                if (cell != null && cell.getCellType() == CellType.STRING)
                    cell.setCellStyle(CustomWorkbook.this.flagStyle);
            }
        }

        /**
         * Auto-size all the columns in the current sheet.
         */
        public void autoSizeColumns() {
            for (int c = 0; c < this.maxCols; c++)
                this.autoSizeColumn(c);
        }


        @Override
        public void close() {
            if (this.tableMode)
                this.makeTable();
            CustomWorkbook.this.openSheets.remove(this);
        }

        /**
         * Convert the cells currently in the sheet to a table.
         */
        private void makeTable() {
            // Delimit the table to the cells created.
            AreaReference range = new AreaReference(ORIGIN_REF,
                    new CellReference(this.rowIdx - 1, this.maxCols - 1), SpreadsheetVersion.EXCEL2007);
            // Create the table.
            XSSFTable myTable = this.sheet.createTable(range);
            // Define the table style.
            CTTable cttable = myTable.getCTTable();
            CTTableStyleInfo tableStyle = cttable.addNewTableStyleInfo();
            tableStyle.setName("TableStyleMedium9");
            tableStyle.setShowColumnStripes(false);
            tableStyle.setShowRowStripes(true);
            // Set up the table name and ID.
            String fixedName = TableName.fix(this.sheet.getSheetName());
            TableName tableIdentifier = CustomWorkbook.this.tableMap.createTable(fixedName);
            cttable.setDisplayName("table_" + fixedName);
            cttable.setName(tableIdentifier.getId());
            cttable.setId(tableIdentifier.getNum());
            // Turn on auto-filter.
            cttable.addNewAutoFilter();
        }

    }

   /**
     * Construct a new, blank workbook to be written to the specified file.
     *
     * @param outFile	workbook output file
     */
    public static CustomWorkbook create(File outFile) {
        CustomWorkbook retVal = new CustomWorkbook();
        retVal.outFile = outFile;
        // Create the workbook.
        retVal.workbook = new XSSFWorkbook();
        retVal.tableMap = new TableNameMap();
        retVal.precision = 2;
        retVal.setupWorkbook();
        return retVal;
    }

    /**
     * Construct a workbook from an existing file.
     *
     * @param inFile	workbook file to update
     *
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static CustomWorkbook load(File inFile) throws InvalidFormatException, IOException {
        CustomWorkbook retVal = new CustomWorkbook();
        retVal.outFile = inFile;
        // Read the workbook.
        try (FileInputStream inStream = new FileInputStream(inFile)) {
            retVal.workbook = new XSSFWorkbook(inStream);
        }
        // Now we need to find the old tables.  We add them to the table map so that new tables will
        // have unique names.
        retVal.tableMap = new TableNameMap();
        var iter = retVal.workbook.sheetIterator();
        while (iter.hasNext()) {
            XSSFSheet currSheet = (XSSFSheet) iter.next();
            for (XSSFTable table : currSheet.getTables()) {
                CTTable cttable = table.getCTTable();
                retVal.tableMap.addTable(cttable.getName(), cttable.getId(), cttable.getDisplayName());
            }
        }
        // Finish setting up the workbook.
        retVal.precision = 2;
        retVal.setupWorkbook();
        return retVal;
    }

    /**
     * Specify a new number precision and update the styles.  This does not change
     * the precision of cells that already exist.
     *
     * @param newPrecision		new number of digits past the decimal
     */
    public void setPrecision(int newPrecision) {
        this.precision = newPrecision;
        this.setupNumberStyles();
    }

    /**
     * Specify a new maximum column width for auto-sizing.
     *
     * @param maxWidth		new maximum width
     */
    public void setMaxWidth(int maxWidth) {
        this.maxWidth = maxWidth;
    }

    /**
     * Perform all the necessary workbook initialization.
     */
    private void setupWorkbook() {
        // Clear the maximum width/
        this.maxWidth = Integer.MAX_VALUE;
        // Denote we have no worksheet.
        this.defaultSheet = null;
        this.openSheets = new ArrayList<>();
        // Denote we have no worksheets to delete or dead tables.
        this.deleteQueue = new ArrayList<>();
        this.deadTables = 0;
        // Set up the creation helper and the formatter.
        this.helper = this.workbook.getCreationHelper();
        this.formatter = this.workbook.createDataFormat();
        short fracFmt = this.formatter.getFormat("#0.0000");
        short intFmt = this.formatter.getFormat("##0");
        short mlFmt = this.formatter.getFormat("#0.0");
        // Create the header style.
        this.headStyle = this.workbook.createCellStyle();
        this.headStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        this.headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // Create the number styles.
        this.intStyle = this.workbook.createCellStyle();
        this.intStyle.setDataFormat(intFmt);
        this.intStyle.setAlignment(HorizontalAlignment.RIGHT);
        this.intStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.fracStyle = this.workbook.createCellStyle();
        this.fracStyle.setDataFormat(fracFmt);
        this.fracStyle.setAlignment(HorizontalAlignment.RIGHT);
        this.fracStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.mlStyle = this.workbook.createCellStyle();
        this.mlStyle.setDataFormat(mlFmt);
        this.mlStyle.setAlignment(HorizontalAlignment.RIGHT);
        this.mlStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.setupNumberStyles();
        // Create the text styles
        this.textStyle = this.workbook.createCellStyle();
        this.textStyle.setAlignment(HorizontalAlignment.LEFT);
        this.textStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.flagStyle = this.workbook.createCellStyle();
        this.flagStyle.setAlignment(HorizontalAlignment.CENTER);
        this.flagStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.wrapStyle = this.workbook.createCellStyle();
        this.wrapStyle.setAlignment(HorizontalAlignment.LEFT);
        this.wrapStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.wrapStyle.setWrapText(true);
        XSSFFont hlinkfont = workbook.createFont();
        hlinkfont.setUnderline(XSSFFont.U_SINGLE);
        hlinkfont.setColor(IndexedColors.INDIGO.getIndex());
        this.linkStyle = this.workbook.createCellStyle();
        this.linkStyle.setAlignment(HorizontalAlignment.LEFT);
        this.linkStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.linkStyle.setFont(hlinkfont);
        this.lwrapStyle = this.workbook.createCellStyle();
        this.lwrapStyle.setAlignment(HorizontalAlignment.LEFT);
        this.lwrapStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.lwrapStyle.setFont(hlinkfont);
        this.lwrapStyle.setWrapText(true);
    }

    /**
     * This method sets up the default-precision number styles.
     */
    private void setupNumberStyles() {
        // Create the level styles.
        short numFmt = this.formatter.getFormat("###0." + StringUtils.repeat('0', this.precision));
        this.numStyle = this.workbook.createCellStyle();
        this.numStyle.setDataFormat(numFmt);
        this.numStyle.setAlignment(HorizontalAlignment.RIGHT);
        this.highStyle = this.workbook.createCellStyle();
        this.highStyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
        this.highStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.highStyle.setDataFormat(numFmt);
        this.highStyle.setAlignment(HorizontalAlignment.RIGHT);
        this.lowStyle = this.workbook.createCellStyle();
        this.lowStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        this.lowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.lowStyle.setDataFormat(numFmt);
        this.lowStyle.setAlignment(HorizontalAlignment.RIGHT);
    }

    /**
     * Create a new worksheet.
     *
     * @param name		name of the new worksheet
     * @param isTable	TRUE to make this worksheet a table
     */
    public void addSheet(String name, boolean isTable) {
        if (this.defaultSheet != null)
            this.closeSheet();
        this.defaultSheet = this.new Sheet(name, isTable);
        this.openSheets.add(this.defaultSheet);
    }

    /**
     * Store the headers for the current worksheet.
     *
     * @param headers	list of header names to store
     */
    public void setHeaders(List<String> headers) {
        this.defaultSheet.setHeaders(headers);
    }

    /**
     * Store a range-colored value in the next cell of the current row.  The value will be normally-colored
     * if it is between the minimum and maximum.  If it is at or below the minimum, it will be red.  If
     * it is at or above the maximum, it will be green.
     *
     * @param value		value to store
     * @param min		maximum "low" value
     * @param max		minimum "high" value
     */
    public void storeCell(double value, double min, double max) {
        this.defaultSheet.storeCell(value, min, max);
    }

    /**
     * Add a new data row to the current sheet.
     */
    public void addRow() {
        this.defaultSheet.addRow();
    }

    /**
     * Store a floating-point value in the next cell of the current row.
     *
     * @param value		value to store
     * @param style		style of number
     */
    public void storeCell(double value, Num style) {
        this.defaultSheet.storeCell(value, style);
    }

    /**
     * Store a floating-point value in a random cell of the current sheet.  Random-access stores should be
     * done after the sequential storing that is the norm.
     *
     * @param r			target row index
     * @param c			target column index
     * @param value		value to store
     * @param style		style of number
     *
     */
    public void storeCell(int r, int c, double value, Num style) {
        this.defaultSheet.storeCell(r, c, value, style);
    }

    /**
     * Store a string value in a random cell.  Random stores should be done after the sequential storing
     * that is the norm.
     *
     * @param r			target row index
     * @param c			target column index
     * @param value		value to store
     */
    public void storeCell(int r, int c, String value) {
        this.defaultSheet.storeCell(r, c, value);
    }

    /**
     * Store an integer value in the next cell of this row.
     *
     * @param value		value to store
     */
    public void storeCell(int value) {
        this.defaultSheet.storeCell(value);
    }

    /**
     * Store a text value in the next cell of this row.
     *
     * @param value		value to store
     * @param style		style of text
     */
    public void storeCell(String value, Text style) {
        this.defaultSheet.storeCell(value, style);
    }

    /**
     * Store a text value and link in the next cell of this row.
     *
     * @param value		value to store
     * @param url		URL for the link (NULL for none)
     * @param comment	comment for the cell (NULL for none)
     */
    public void storeCell(String value, String url, String comment) {
        this.defaultSheet.storeCell(value, url, comment);
    }

    /**
     * Store an empty cell in the current position.
     */
    public void storeBlankCell() {
        this.defaultSheet.storeBlankCell();
    }

    /**
     * Store an integer cell with a hyperlink.
     *
     * @param value		integer value in the cell
     * @param url		URL for the link (or NULL if none)
     * @param comment	comment text (or NULL if no comment)
     */
    public void storeCell(int value, String url, String comment) {
        this.defaultSheet.storeCell(value, url, comment);
    }

    /**
     * Store a string in a cell and format it normally.
     *
     * @param value		string to store
     */
    public void storeCell(String string) {
        this.defaultSheet.storeCell(string);
    }

    /**
     * Store a number in a cell and format it normally.
     *
     * @param value		number to store
     */
    public void storeCell(double value) {
        this.defaultSheet.storeCell(value);
    }

    /**
     * Autosize the specified column.
     *
     * @param c		index of the column to autosize
     */
    public void autoSizeColumn(int c) {
        this.defaultSheet.autoSizeColumn(c);
    }

    /**
     * Reformat a text column as flags.
     *
     * @param c		index of column to reformat
     */
    public void reformatFlagColumn(int c) {
        this.defaultSheet.reformatFlagColumn(c);
    }

    /**
     * Reformat a number column as integer.
     *
     * @param c		index of the column to reformat
     */
    public void reformatIntColumn(int c) {
        this.defaultSheet.reformatIntColumn(c);
    }

    /**
     * Auto-size all the columns in the current sheet.
     */
    public void autoSizeColumns() {
        this.defaultSheet.autoSizeColumns();
    }

    /**
     * Finalize the current sheet.
     */
    public void closeSheet() {
        // Currently, we just need to convert it to a table if this is table mode.
        if (this.defaultSheet != null)
            this.defaultSheet.close();
        // Insure we don't close it again.
        this.defaultSheet = null;
    }

    @Override
    public void close() {
        // Insure the open sheets are closed.
        while (! this.openSheets.isEmpty())
            this.openSheets.get(0).close();
        this.defaultSheet = null;
        // Delete the dead sheets.
        for (XSSFSheet deadSheet : this.deleteQueue) {
            int deadIdx = this.workbook.getSheetIndex(deadSheet);
            this.workbook.removeSheetAt(deadIdx);
        }
        // Here we write out the Excel file, de-checking any IO exception that occurs.
        try (OutputStream outStream = new FileOutputStream(this.outFile)) {
            this.workbook.write(outStream);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

}
