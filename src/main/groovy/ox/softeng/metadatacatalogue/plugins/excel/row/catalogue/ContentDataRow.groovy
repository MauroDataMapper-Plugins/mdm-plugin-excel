package ox.softeng.metadatacatalogue.plugins.excel.row.catalogue

import ox.softeng.metadatacatalogue.core.catalogue.CatalogueItem
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElement
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.DataType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.ReferenceType
import ox.softeng.metadatacatalogue.plugins.excel.row.StandardDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.column.MetadataColumn

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATACLASS_PATH_SPLIT_REGEX

/**
 * @since 01/03/2018
 */
class ContentDataRow extends StandardDataRow<EnumerationValueDataRow> {

    String dataClassPath
    String dataElementName
    String description
    Integer minMultiplicity
    Integer maxMultiplicity
    String dataTypeName
    String dataTypeDescription
    String referenceToDataClassPath


    ContentDataRow() {
        super(EnumerationValueDataRow)
    }

    ContentDataRow(CatalogueItem catalogueItem) {
        this()

        description = catalogueItem.description

        catalogueItem.metadata.each {md ->
            metadata += new MetadataColumn(namespace: md.namespace, key: md.key, value: md.value)
        }

        if (catalogueItem.instanceOf(DataClass)) {
            dataClassPath = buildPath(catalogueItem)
            minMultiplicity = catalogueItem.minMultiplicity
            maxMultiplicity = catalogueItem.maxMultiplicity
        } else if (catalogueItem.instanceOf(DataElement)) {
            dataClassPath = buildPath(catalogueItem.dataClass)
            dataElementName = catalogueItem.label
            minMultiplicity = catalogueItem.minMultiplicity
            maxMultiplicity = catalogueItem.maxMultiplicity

            DataType dataType = catalogueItem.getDataType()
            dataTypeName = dataType.getLabel()
            dataTypeDescription = dataType.getDescription()

            if (dataType.instanceOf(ReferenceType)) {
                referenceToDataClassPath = buildPath(dataType.getReferenceClass())
            } else if (dataType.instanceOf(EnumerationType)) {
                dataType.getEnumerationValues().sort {it.key}.each {ev ->
                    mergedContentRows += new EnumerationValueDataRow(ev)
                }
            }
        }
    }

    @Override
    Integer getFirstMetadataColumn() {
        10
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)

        dataClassPath = getCellValueAsString(row, 0)
        dataElementName = getCellValueAsString(row, 1)
        description = getCellValueAsString(row, 2)
        minMultiplicity = getCellValueAsInteger(row.getCell(3), 1)
        maxMultiplicity = getCellValueAsInteger(row.getCell(4), 1)
        dataTypeName = getCellValueAsString(row, 5)
        dataTypeDescription = getCellValueAsString(row, 6)
        referenceToDataClassPath = getCellValueAsString(row, 7)

        extractMetadataFromColumnIndex()
    }

    List<String> getDataClassPathList() {
        dataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    List<String> getReferenceToDataClassPathList() {
        referenceToDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    String buildPath(DataClass dataClass) {
        if (!dataClass.parentDataClass) return dataClass.label
        "${buildPath(dataClass.parentDataClass)}|${dataClass.label}"
    }

    Row buildRow(Row row) {
        addCellToRow(row, 0, dataClassPath)
        addCellToRow(row, 1, dataElementName)
        addCellToRow(row, 2, description, true)
        addCellToRow(row, 3, minMultiplicity)
        addCellToRow(row, 4, maxMultiplicity)
        addCellToRow(row, 5, dataTypeName)
        addCellToRow(row, 6, dataTypeDescription, true)
        addCellToRow(row, 7, referenceToDataClassPath)

        addMetadataToRow(row)

        if (mergedContentRows) {
            Row lastRow = row
            Sheet sheet = row.sheet
            for (int r = 0; r < mergedContentRows.size(); r++) {
                EnumerationValueDataRow mcr = mergedContentRows[r]
                lastRow = CellUtil.getRow(row.rowNum + r, sheet)
                mcr.buildRow(lastRow)
            }
            if (row.rowNum != lastRow.rowNum) {
                for (int i = 0; i < sheet.getRow(METADATA_KEY_ROW).lastCellNum; i++) {
                    if (i == 8 || i == 9) continue // These are the enum columns
                    sheet.addMergedRegion(new CellRangeAddress(row.rowNum, lastRow.rowNum, i, i))
                }
            }
        }
        row
    }
}
