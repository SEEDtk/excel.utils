package org.theseed.excel;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;
import org.theseed.io.TabbedLineReader;

/**
 * This creates a custom workbook.  Sadly, the user has to look at the file to figure out if it is
 * any good.
 *
 * @author Bruce Parrello
 *
 */
public class TestCustomWorkbook {

    @Test
    public void testCustomWorkbook() throws IOException, InvalidFormatException {
        File inFile = new File("data", "test.tbl");
        File inFile2 = new File("data", "test2.tbl");
        File outFile = new File("data", "test.xlsx");
        File testFile = new File("data", "test2.xlsx");
        try (CustomWorkbook workbook = CustomWorkbook.create(outFile)) {
            workbook.addSheet("test sheet", true);
            createTestSheet(inFile, workbook);
            workbook.addSheet("norm sheet", false);
            createTestSheet(inFile2, workbook);
            workbook.addSheet("new sheet", true);
            createTestSheet(inFile, workbook);
        }
        FileUtils.copyFile(outFile, testFile);
        try (CustomWorkbook workbook = CustomWorkbook.load(testFile)) {
            workbook.addSheet("newer sheet", false);
            createTestSheet(inFile, workbook);
            workbook.addSheet("test sheet", true);
            createTestSheet(inFile2, workbook);
        }
    }

    private void createTestSheet(File inFile, CustomWorkbook workbook) throws IOException {
        workbook.setHeaders(Arrays.asList("Fid", "gene", "val1", "val2", "val3", "notes", "thing"));
        try (TabbedLineReader inStream = new TabbedLineReader(inFile)) {
            for (TabbedLineReader.Line line : inStream) {
                workbook.addRow();
                workbook.storeCell(line.get(0), CustomWorkbook.Text.NORMAL);
                workbook.storeCell(line.get(1), "https://rnaseq.theseed.org/" + line.get(1), "Comment for " + line.get(1));
                workbook.storeCell(line.getInt(2));
                workbook.storeCell(line.getDouble(3), CustomWorkbook.Num.FRACTION);
                workbook.storeCell(line.getDouble(4), 150.0, 250.0);
                workbook.storeCell(line.get(5), CustomWorkbook.Text.FLAG);
                workbook.storeCell(line.get(6), CustomWorkbook.Text.NORMAL);
            }
        }
    }

}
