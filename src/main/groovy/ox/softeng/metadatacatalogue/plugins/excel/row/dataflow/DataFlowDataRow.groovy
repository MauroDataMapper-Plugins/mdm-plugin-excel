package ox.softeng.metadatacatalogue.plugins.excel.row.dataflow

import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlow
import ox.softeng.metadatacatalogue.plugins.excel.row.StandardDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.column.MetadataColumn

import org.apache.poi.ss.usermodel.Row

/**
 * @since 09/03/2018
 */
class DataFlowDataRow extends StandardDataRow {

    String sheetKey
    String name
    String description
    String sourceDataModelName
    String targetDataModelName

    DataFlow dataFlow

    DataFlowDataRow() {
        super()
    }

    DataFlowDataRow(DataFlow dataFlow) {
        this()
        this.dataFlow = dataFlow
        sheetKey = ''
        String[] words = dataFlow.label.split(' ')
        if (words.size() == 1) {
            sheetKey = dataFlow.label.toUpperCase()
        } else {
            words.each {word -> sheetKey += word[0].toUpperCase()}
        }
        name = dataFlow.label
        description = dataFlow.description
        sourceDataModelName = dataFlow.source.label
        targetDataModelName = dataFlow.target.label

        dataFlow.metadata.each {md ->
            metadata += new MetadataColumn(namespace: md.namespace, key: md.key, value: md.value)
        }
    }

    @Override
    Integer getFirstMetadataColumn() {
        5
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)
        sheetKey = getCellValueAsString(row, 0)
        name = getCellValueAsString(row, 1)
        description = getCellValueAsString(row, 2)
        sourceDataModelName = getCellValueAsString(row, 3)
        targetDataModelName = getCellValueAsString(row, 4)

        extractMetadataFromColumnIndex()
    }

    Row buildRow(Row row) {
        addCellToRow(row, 0, sheetKey)
        addCellToRow(row, 1, name)
        addCellToRow(row, 2, description, true)
        addCellToRow(row, 3, sourceDataModelName)
        addCellToRow(row, 4, targetDataModelName)

        addMetadataToRow(row)
        row
    }
}
