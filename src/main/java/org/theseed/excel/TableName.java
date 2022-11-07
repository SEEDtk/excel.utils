/**
 *
 */
package org.theseed.excel;

import org.apache.commons.lang3.StringUtils;
import org.theseed.magic.MagicObject;

/**
 * This is a magic object used to generate table identifiers.  There is no normalization, since table identifiers
 * are computed from the sheet name, and the sheet name is unique.  The goal is to produce a safe table name
 * from the sheet name.
 *
 * @author Bruce Parrello
 *
 */
public class TableName extends MagicObject {

    // FIELDS
    /** serialization object version */
    private static final long serialVersionUID = -3472320204994842091L;
    /** table ID number */
    private long tableNum;

    /**
     * Create a blank, empty table name object.
     */
    public TableName() {
        this.tableNum = 0;
    }

    /**
     * Create a table name object from an existing table.
     *
     * @param id	table identifier (or NULL if one needs to be created)
     * @param num	table number
     * @param name	sheet name containing the table
     */
    public TableName(String id, long num, String name) {
        super(id, name);
        this.tableNum = num;
    }

    /**
     * Convert a sheet name string to an acceptable format for a table name.  This involves
     * (1) removing the word "sheet" at the end, (2) converting everything to lower case, and (3) turning
     * all illegal characters to underscores.
     *
     * @param input		string to convert
     *
     * @return a clean string
     */
    public static String fix(String input) {
        String retVal;
        if (StringUtils.isBlank(input))
            retVal = "_";
        else {
            retVal = input.replaceAll("[\\W_]+", "_").toLowerCase();
            retVal = StringUtils.removeEnd(retVal, "_sheet");
        }
        return retVal;
    }

    /**
     * @return the table number
     */
    public long getNum() {
        return this.tableNum;
    }

    @Override
    protected String normalize(String name) {
        return name;
    }

}
