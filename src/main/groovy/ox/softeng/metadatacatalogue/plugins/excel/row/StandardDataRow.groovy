package ox.softeng.metadatacatalogue.plugins.excel.row

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationType

import org.apache.poi.ss.usermodel.Row

/**
 * @since 07/09/2017
 */
abstract class StandardDataRow<K extends EnumerationDataRow> extends DataRow {

    Class<K> enumerationExcelDataRowClass
    List<K> mergedContentRows

    StandardDataRow() {
        super()
        mergedContentRows = []
    }

    StandardDataRow(Class<K> enumerationExcelDataRowClass) {
        this()
        this.enumerationExcelDataRowClass = enumerationExcelDataRowClass
    }

    void addToMergedContentRows(Row row) {
        if (enumerationExcelDataRowClass) {
            K enumerationRow = enumerationExcelDataRowClass.newInstance()
            enumerationRow.setAndInitialise(row)
            mergedContentRows += enumerationRow
        }
    }

    boolean matchesEnumerationType(EnumerationType enumerationType) {
        if (!mergedContentRows) return false

        if (enumerationType.enumerationValues.size() != mergedContentRows.size()) return false

        // Check every enumeration value has an entry in the merged content rows
        enumerationType.enumerationValues.every {ev ->
            mergedContentRows.any {it.key == ev.key && it.value == ev.value}
        }
    }
}
