package ox.softeng.metadatacatalogue.plugins.excel.row.dataflow

import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowComponent
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.plugins.excel.row.StandardDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.column.MetadataColumn

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATACLASS_PATH_SPLIT_REGEX

/**
 * @since 09/03/2018
 */
class DataFlowComponentDataRow extends StandardDataRow<ElementDataRow> {

    String sourceDataClassPath
    String sourceDataElementName
    String transformLabel
    String transformDescription
    String targetDataElementName
    String targetDataClassPath

    DataFlowComponentDataRow() {
        super(ElementDataRow)
    }

    DataFlowComponentDataRow(DataFlowComponent dataFlowComponent) {
        this()
        transformLabel = dataFlowComponent.label
        transformDescription = dataFlowComponent.description

        for (int i = 0; i < Math.max(dataFlowComponent.sourceElements.size(), dataFlowComponent.targetElements.size()); i++) {
            mergedContentRows += new ElementDataRow(dataFlowComponent.sourceElements[i], dataFlowComponent.targetElements[i])
        }

        dataFlowComponent.metadata.each {md ->
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

        sourceDataClassPath = getCellValueAsString(row, 0)
        sourceDataElementName = getCellValueAsString(row, 1)

        transformLabel = getCellValueAsString(row, 2)
        transformDescription = getCellValueAsString(row, 3)

        targetDataElementName = getCellValueAsString(row, 4)
        targetDataClassPath = getCellValueAsString(row, 5)

        extractMetadataFromColumnIndex()
    }

    List<String> getSourceDataClassPathList() {
        sourceDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    List<String> getTargetDataClassPathList() {
        targetDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    String buildPath(DataClass dataClass) {
        if (!dataClass.parentDataClass) return dataClass.label
        "${buildPath(dataClass.parentDataClass)}|${dataClass.label}"
    }

    int findFirstRowWithNoData(Sheet sheet, int start, int finish, int col) {
        for (int i = start; i <= finish; i++) {
            if (!getCellValueAsString(sheet.getRow(i), col)) return i
        }
        -1
    }

    Row buildRow(Row row) {
        addCellToRow(row, 2, transformLabel)
        addCellToRow(row, 3, transformDescription, true)

        addMetadataToRow(row)

        if (mergedContentRows) {
            Row lastRow = row
            Sheet sheet = row.sheet
            for (int r = 0; r < mergedContentRows.size(); r++) {
                ElementDataRow edr = mergedContentRows[r]
                lastRow = CellUtil.getRow(row.rowNum + r, sheet)
                edr.buildRow(lastRow)
            }
            if (row.rowNum != lastRow.rowNum) {
                for (int i = 0; i < sheet.getRow(METADATA_KEY_ROW).lastCellNum; i++) {
                    if (i in [0, 1, 4, 5]) {
                        int r = findFirstRowWithNoData(sheet, row.rowNum, lastRow.rowNum, i)
                        if (r != -1) {
                            sheet.addMergedRegion(new CellRangeAddress(r - 1, lastRow.rowNum, i, i))
                        }
                    } else sheet.addMergedRegion(new CellRangeAddress(row.rowNum, lastRow.rowNum, i, i))
                }
            }
        }
        row
    }
}
