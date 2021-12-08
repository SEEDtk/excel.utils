/**
 *
 */
package org.theseed.excel;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

import org.junit.jupiter.api.Test;

/**
 * @author Bruce Parrello
 *
 */
public class TestTableName {

    @Test
    public void testNameFix() {
        String[] original = new String[] { "test sheet", "New, improved, deluxe sheet", "Goofy34!@#$%^()_characters__ in %%%%line",
                "Hello, young lovers", "", null };
        String[] improved = new String[] { "test", "new_improved_deluxe", "goofy34_characters_in_line", "hello_young_lovers",
                "_", "_" };
        for (int i = 0; i < original.length; i++) {
            assertThat(Integer.toString(i), TableName.fix(original[i]), equalTo(improved[i]));
        }
    }


}
