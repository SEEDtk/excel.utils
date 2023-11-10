/**
 *
 */
package org.theseed.excel.utils;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

import java.io.File;
import java.util.Arrays;

import org.junit.jupiter.api.Test;
import org.theseed.basic.ParseFailureException;

/**
 *
 * @author Bruce Parrello
 *
 */
class TestDistributor {

    public static double[] series1 = new double[] { 0.01, 0.02, 0.03, 0.11, 0.12, 0.13, 0.14, 0.21, 0.22, 0.23, 0.24, 0.25,
            0.31, 0.32, 0.33, 0.34, 0.35, 0.36, 0.71, 0.72, 0.73, 0.74, 0.81, 0.82, 0.91 };
    public static double[] series2 = new double[] { 0.0, 0.11, 0.21, 0.31, 0.41, 0.51, 0.61, 0.71, 0.81, 0.91, 1.0 };
    public static int[] series1Buckets = new int[] { 3, 4, 5, 6, 0, 0, 0, 4, 2, 1 };
    public static int[] series2Buckets = new int[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 2 };

    @Test
    void test() throws ParseFailureException {
        Distributor bucketMap = new Distributor(0.0, 1.0, 10);
        Arrays.stream(series1).forEach(x -> bucketMap.addValue("series1", x));
        Arrays.stream(series2).forEach(x -> bucketMap.addValue("series2", x));
        assertThat(bucketMap.getBuckets("series3"), nullValue());
        assertThat(bucketMap.getBuckets("series2"), equalTo(series2Buckets));
        assertThat(bucketMap.getBuckets("series1"), equalTo(series1Buckets));
        File outFile = new File("data", "buckets.xlsx");
        bucketMap.save(outFile);
    }

}
