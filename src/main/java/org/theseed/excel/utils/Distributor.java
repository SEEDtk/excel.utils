/**
 *
 */
package org.theseed.excel.utils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.theseed.basic.ParseFailureException;
import org.theseed.excel.CustomWorkbook;

/**
 * This object creates a distribution spreadsheet.  The constructor passes in a minimum, a maximum, and a number
 * of buckets.  Each value passed in has a series name and a value between the extremes.  We count the value in
 * the appropriate bucket.  A method is provided to write the results to an Excel spreadsheet in the form of a
 * table.  Each row is a bucket and each column is a series name.
 *
 * @author Bruce Parrello
 *
 */
public class Distributor {

    // FIELDS
    /** logging facility */
    protected static Logger log = LoggerFactory.getLogger(Distributor.class);
    /** map of series names to bucket arrays */
    private Map<String, int[]> bucketMap;
    /** bucket size */
    private double bucketWidth;
    /** size of each bucket array */
    private int nBuckets;
    /** minimum value */
    private double minimum;
    /** recommended display precision for floating-point */
    private int precision;

    /**
     * Construct a distributor.
     *
     * @param min		minimum possible value
     * @param max		maximum possible value
     * @param n			number of desired buckets
     *
     * @throws ParseFailureException
     */
    public Distributor(double min, double max, int n) throws ParseFailureException {
        this.nBuckets = n;
        if (n <= 1)
            throw new ParseFailureException("Cannot create a distribution with less than 2 buckets.");
        if (min >= max)
            throw new ParseFailureException("Minimum of range must be less than maximum.");
        this.minimum = min;
        this.bucketWidth = (max - min) / n;
        // We use a tree map so that the series names are sorted, and because we expect the series count to be small.
        this.bucketMap = new TreeMap<String, int[]>();
        // Compute the recommended format.
        int digits = (int) Math.ceil(Math.log10(Math.abs(max))) + 1;
        int divisor = (int) Math.ceil(Math.log10(Math.abs(n))) + 1;
        this.precision = 0;
        if (digits - divisor < 0) this.precision = divisor - digits;
    }

    /**
     * Add a value to a series.
     *
     * @param name		name of the series
     * @param value		value to record
     */
    public void addValue(String name, double value) {
        this.addValues(name, value);
    }

    /**
     * Add an array of values to a series.
     *
     * @param name		name of the series
     * @param values	values to record
     */
    public void addValues(String name, double... values) {
        // We count on Java's habit of initializing all ints to zero.
        int[] buckets = this.bucketMap.computeIfAbsent(name, x -> new int[this.nBuckets]);
        for (double value : values) {
            int idx = (int) ((value - this.minimum) / this.bucketWidth);
            // Catch the maximum if it happens.
            if (idx >= this.nBuckets) idx = this.nBuckets - 1;
            // Count the value.
            buckets[idx]++;
        }
    }

    /**
     * Get the bucket array for a series.
     *
     * @param name		name of the series
     *
     * @return the distribution counts for the named series, or NULL if it does not exist
     */
    public int[] getBuckets(String name) {
        return this.bucketMap.get(name);
    }

    /**
     * @return the minimum value for a bucket
     *
     * @param idx	bucket index
     */
    protected double getLower(int idx) {
        return this.minimum + idx * this.bucketWidth;
    }

    /**
     * Save a spreadsheet for this distribution.
     *
     * @param outFile	name of the file in which to store the spreadsheet
     */
    public void save(File outFile) {
        try (CustomWorkbook workbook = CustomWorkbook.create(outFile)) {
            log.info("Saving distribution data to {}.", outFile);
            workbook.addSheet("Distribution", true);
            workbook.setPrecision(this.precision);
            // Get an ordered list of the buckets.
            var names = this.bucketMap.keySet();
            var bucketList = names.stream().map(x -> this.bucketMap.get(x)).collect(Collectors.toList());
            // Create the header list.
            List<String> headers = new ArrayList<String>(nBuckets + 1);
            headers.add("bucket_min");
            headers.addAll(names);
            workbook.setHeaders(headers);
            // Now loop through the rows (one per bucket), filling in the cells.
            for (int idx = 0; idx < this.nBuckets; idx++) {
                workbook.addRow();
                workbook.storeCell(this.getLower(idx));
                for (int[] buckets : bucketList)
                    workbook.storeCell(buckets[idx]);
            }
        }
    }

}
