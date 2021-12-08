/**
 *
 */
package org.theseed.excel;

import org.theseed.magic.MagicMap;

/**
 * @author Bruce Parrello
 *
 */
public class TableNameMap extends MagicMap<TableName> {

    // FIELDS
    /** last table number used */
    private long tableIdx;

    /**
     * Construct a new table-name map.
     */
    public TableNameMap() {
        super(new TableName());
        this.tableIdx = 0;
    }

    /**
     * Add the identifying information for an existing table.
     */
    public TableName addTable(String id, long num, String name) {
        TableName retVal = new TableName(id, num, name);
        if (num > this.tableIdx)
            this.tableIdx = num;
        this.register(retVal);
        return retVal;
    }

    /**
     * Add a new table.
     *
     * @param name		name of sheet containing the table
     *
     * @return the identifier object for the new table
     */
    public TableName createTable(String name) {
        // Get the table number, which must be unique.
        this.tableIdx++;
        long num = this.tableIdx;
        // Does this table already exist? If so, we just return it.
        TableName retVal = this.getByName(name);
        if (retVal == null) {
            // Create a table name with a blank ID.  The ID will generate when we put it in the map.
            retVal = new TableName(null, num, name);
            this.put(retVal);
        }
        // Return the name object with the generated ID and number inside.
        return retVal;
    }

}
