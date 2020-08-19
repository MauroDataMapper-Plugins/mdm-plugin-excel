package ox.softeng.metadatacatalogue.plugins.excel.row.catalogue

import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.plugins.excel.row.StandardDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.column.MetadataColumn

import org.apache.poi.ss.usermodel.Row

/**
 * @since 01/03/2018
 */
class DataModelDataRow extends StandardDataRow {

    String sheetKey
    String name
    String description
    String author
    String organisation
    String type

    DataModel dataModel

    DataModelDataRow() {
        super()
    }

    DataModelDataRow(DataModel dataModel) {
        this()
        this.dataModel = dataModel
        sheetKey = ''
        String[] words = dataModel.label.split(' ')
        if (words.size() == 1) {
            sheetKey = dataModel.label.toUpperCase()
        } else {
            words.each {word ->
                if(word.length() > 0) {
                    sheetKey += word[0].toUpperCase()
                }
            }
        }
        name = dataModel.label
        description = dataModel.description
        author = dataModel.author
        organisation = dataModel.organisation
        type = dataModel.type.label

        dataModel.metadata.each {md ->
            metadata += new MetadataColumn(namespace: md.namespace, key: md.key, value: md.value)
        }
    }

    @Override
    Integer getFirstMetadataColumn() {
        6
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)
        sheetKey = getCellValueAsString(row, 0)
        name = getCellValueAsString(row, 1)
        description = getCellValueAsString(row, 2)
        author = getCellValueAsString(row, 3)
        organisation = getCellValueAsString(row, 4)
        type = getCellValueAsString(row, 5)

        extractMetadataFromColumnIndex()
    }

    Row buildRow(Row row) {
        addCellToRow(row, 0, sheetKey)
        addCellToRow(row, 1, name)
        addCellToRow(row, 2, description, true)
        addCellToRow(row, 3, author)
        addCellToRow(row, 4, organisation)
        addCellToRow(row, 5, type)

        addMetadataToRow(row)
        row
    }
}
