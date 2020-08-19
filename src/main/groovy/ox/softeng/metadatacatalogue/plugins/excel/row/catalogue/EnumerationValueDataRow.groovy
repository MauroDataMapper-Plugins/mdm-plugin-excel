package ox.softeng.metadatacatalogue.plugins.excel.row.catalogue

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationValue
import ox.softeng.metadatacatalogue.plugins.excel.row.EnumerationDataRow

import org.apache.poi.ss.usermodel.Row

/**
 * @since 01/03/2018
 */
class EnumerationValueDataRow extends EnumerationDataRow {

    public static final int KEY_COL_INDEX = 8
    public static final int VALUE_COL_INDEX = 9

    EnumerationValueDataRow() {
        super()
    }

    EnumerationValueDataRow(EnumerationValue enumerationValue) {
        this()
        key = enumerationValue.key
        value = enumerationValue.value
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)
        key = getCellValueAsString(row, KEY_COL_INDEX)
        value = getCellValueAsString(row, VALUE_COL_INDEX)
    }

    Row buildRow(Row row) {
        addCellToRow(row, KEY_COL_INDEX, key)
        addCellToRow(row, VALUE_COL_INDEX, value)
        row
    }
}
