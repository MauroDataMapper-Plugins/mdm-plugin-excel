/*
 * Copyright 2020-2023 University of Oxford and NHS England
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

import groovy.transform.CompileDynamic
import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

@Slf4j
@CompileStatic
abstract class DataRow implements CellHandler {

    static final int METADATA_NAMESPACE_ROW_INDEX = 0
    static final int METADATA_KEY_ROW_INDEX = 1

    Row row
    private List<MetadataColumn> metadata = []

    @CompileDynamic
    static Map<String, Set<String>> getMetadataNamespaceAndKeys(List<DataRow> dataRows) {
        dataRows.collect { it.metadata }.flatten().groupBy { it.namespace }
                .collectEntries { [it.key, it.value.key as Set] }
                .findAll { it.value } as Map<String, Set<String>>
    }

    List<MetadataColumn> getMetadata() {
        return metadata
    }

    abstract void initialiseRow(Row row)

    abstract Row buildRow(Row row)

    abstract int getFirstMetadataColumn()

    @CompileDynamic
    @SuppressWarnings('GroovyAssignabilityCheck')
    void addCellToRow(Row row, int cellNum, Object content, boolean wrapText = false) {
        Cell cell = row.createCell(cellNum)
        cell.setCellValue(content)
        if (wrapText) {
            CellUtil.setCellStyleProperty(cell, CellUtil.WRAP_TEXT, wrapText)
            cell.cellStyle.setShrinkToFit(!wrapText)
        }
    }

    void addMetadataToRow(Row row) {
        if (!metadata) return
        Row namespaceRow = row.sheet.getRow(METADATA_NAMESPACE_ROW_INDEX)
        Row keyRow = row.sheet.getRow(METADATA_KEY_ROW_INDEX)
        metadata.groupBy { it.namespace }.each { String namespace, List<MetadataColumn> metadataColumns ->
            Cell namespaceHeader = namespaceRow.find { Cell cell -> getCellValue(cell) == namespace } as Cell
            CellRangeAddress mergeRegion = getMergeRegion(namespaceHeader)
            metadataColumns.each { MetadataColumn metadata ->
                log.debug('Adding {}:{}', metadata.namespace, metadata.key)
                Cell keyHeader = mergeRegion
                    ? findCell(keyRow, metadata.key, mergeRegion.firstColumn, mergeRegion.lastColumn)
                    : findCell(keyRow, metadata.key, namespaceHeader.columnIndex, namespaceHeader.columnIndex)
                addCellToRow(row, keyHeader.columnIndex, metadata.value, true)
            }
        }
    }

    void addToMetadata(String namespace, String key, String value) {
        metadata << new MetadataColumn(namespace: namespace, key: key, value: value)
    }

    void extractMetadataFromColumnIndex() {
        Row namespaceRow = row.sheet.getRow(METADATA_NAMESPACE_ROW_INDEX)
        Row keyRow = row.sheet.getRow(METADATA_KEY_ROW_INDEX)
        List<Cell> cells = row.findAll { Cell cell ->
            cell.columnIndex >= firstMetadataColumn &&
            cell.columnIndex <= Math.max(namespaceRow.lastCellNum, keyRow.lastCellNum) &&
            getCellValue(cell)
        } as List<Cell>
        cells.each { Cell cell ->
            CellRangeAddress mergeRegion = getMergeRegion(row.sheet, namespaceRow.rowNum, cell.columnIndex)
            String namespaceHeader = getCellValue(namespaceRow, cell.columnIndex)
            String keyHeader = getCellValue(keyRow, cell.columnIndex)
            String namespace = keyHeader ? namespaceHeader ?: getCellValue(namespaceRow, mergeRegion?.firstColumn) : null
            String key = keyHeader ?: namespaceHeader
            metadata << new MetadataColumn(namespace: namespace, key: key, value: getCellValue(cell))
        }
    }

    private Cell findCell(Row row, String content, int firstColumn, int lastColumn) {
        row.find { Cell cell -> cell.columnIndex >= firstColumn && cell.columnIndex <= lastColumn && getCellValue(cell) == content } as Cell
    }
}
