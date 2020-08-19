/*
 * Copyright 2020 University of Oxford
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * SPDX-License-Identifier: Apache-2.0
 */
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.catalogue

import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataType
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.StandardDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column.MetadataColumn

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

import static uk.ac.ox.softeng.maurodatamapper.plugins.excel.ExcelPlugin.DATACLASS_PATH_SPLIT_REGEX

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
