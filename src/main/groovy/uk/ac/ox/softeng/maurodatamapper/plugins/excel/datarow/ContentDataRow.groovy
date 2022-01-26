/*
 * Copyright 2020-2022 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow

import uk.ac.ox.softeng.maurodatamapper.core.facet.Metadata
import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.enumeration.EnumerationValue

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

@CompileStatic
class ContentDataRow extends StandardDataRow<EnumerationDataRow> {

    private static final String DATACLASS_PATH_SPLIT_REGEX = ~/\|/

    String dataClassPath
    String dataElementName
    String description
    String dataTypeName
    String dataTypeDescription
    String referenceToDataClassPath

    int minMultiplicity
    int maxMultiplicity

    ContentDataRow() {
        super(EnumerationDataRow)
    }

    ContentDataRow(CatalogueItem catalogueItem) {
        description = catalogueItem.description

        if (catalogueItem instanceof DataClass) {
            dataClassPath = buildPath(catalogueItem)
            minMultiplicity = catalogueItem.minMultiplicity ?: 0
            maxMultiplicity = catalogueItem.maxMultiplicity ?: 0
        } else if (catalogueItem instanceof DataElement) {
            dataClassPath = buildPath(catalogueItem.dataClass)
            dataElementName = catalogueItem.label
            minMultiplicity = catalogueItem.minMultiplicity ?: 0
            maxMultiplicity = catalogueItem.maxMultiplicity ?: 0

            DataType dataType = catalogueItem.dataType
            dataTypeName = dataType.label
            dataTypeDescription = dataType.description
            if (dataType instanceof ReferenceType) {
                referenceToDataClassPath = buildPath(dataType.referenceClass)
            } else if (dataType instanceof EnumerationType) {
                dataType.enumerationValues.sort { it.key }.each { EnumerationValue enumerationValue ->
                    mergedContentRows << new EnumerationDataRow(enumerationValue)
                }
            }
        }
    }

    @Override
    void initialiseRow(Row row) {
        setRow(row)
        dataClassPath = getCellValue(row, 0)
        dataElementName = getCellValue(row, 1)
        description = getCellValue(row, 2)
        minMultiplicity = getCellValueAsInt(row.getCell(3), 1)
        maxMultiplicity = getCellValueAsInt(row.getCell(4), 1)
        dataTypeName = getCellValue(row, 5)
        dataTypeDescription = getCellValue(row, 6)
        referenceToDataClassPath = getCellValue(row, 7)
        extractMetadataFromColumnIndex()
    }

    @Override
    Row buildRow(Row row) {
        addCellToRow row, 0, dataClassPath
        addCellToRow row, 1, dataElementName
        addCellToRow row, 2, description, true
        addCellToRow row, 3, minMultiplicity
        addCellToRow row, 4, maxMultiplicity
        addCellToRow row, 5, dataTypeName
        addCellToRow row, 6, dataTypeDescription, true
        addCellToRow row, 7, referenceToDataClassPath
        addMetadataToRow row

        if (!mergedContentRows) return row

        Row lastRow = row
        Sheet sheet = row.sheet
        mergedContentRows.size().times { int rowNum ->
            lastRow = CellUtil.getRow(row.rowNum + rowNum, sheet)
            mergedContentRows[rowNum].buildRow(lastRow)
        }
        if (row.rowNum != lastRow.rowNum) {
            sheet.getRow(METADATA_KEY_ROW_INDEX).lastCellNum.times { int columnIndex ->
                if (columnIndex == 8 || columnIndex == 9) return // Enum columns
                sheet.addMergedRegion(new CellRangeAddress(row.rowNum, lastRow.rowNum, columnIndex, columnIndex))
            }
        }

        row
    }

    @Override
    int getFirstMetadataColumn() {
        10
    }

    List<String> getDataClassPathList() {
        dataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    List<String> getReferenceToDataClassPathList() {
        referenceToDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    private int getCellValueAsInt(Cell cell, int defaultValue) {
        String cellValue = getCellValue(cell)
        cellValue ? cellValue.toInteger() : defaultValue
    }

    private String buildPath(DataClass dataClass) {
        if (!dataClass.parentDataClass) return dataClass.label
        "${buildPath(dataClass.parentDataClass)}|${dataClass.label}"
    }
}
