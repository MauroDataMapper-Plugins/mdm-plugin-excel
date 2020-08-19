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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.row

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column.MetadataColumn
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.util.CellHandler

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.slf4j.Logger
import org.slf4j.LoggerFactory

/**
 * @since 02/03/2018
 */
abstract class DataRow implements CellHandler {

    public static final Integer METADATA_NAMESPACE_ROW = 0
    public static final Integer METADATA_KEY_ROW = 1

    Row row

    List<MetadataColumn> metadata

    DataRow() {
        metadata = []
    }

    Integer getFirstMetadataColumn() {
        null
    }

    abstract void setAndInitialise(Row row)

    Logger getLogger() {
        LoggerFactory.getLogger(getClass())
    }

    Row buildRow(Row row) {
        row
    }

    void extractMetadataFromColumnIndex() {
        Row firstRow = row.sheet.getRow(METADATA_NAMESPACE_ROW)
        Row secondRow = row.sheet.getRow(METADATA_KEY_ROW)

        int lastCellNum = Math.max(firstRow.lastCellNum, secondRow.lastCellNum)

        List<Cell> cells = row.findAll {it.columnIndex >= firstMetadataColumn && it.columnIndex <= lastCellNum && getCellValueAsString(it)}

        cells.each {
            int columnIndex = it.columnIndex
            CellRangeAddress mergeRegion = getMergeRegion(row.sheet, firstRow.rowNum, columnIndex)

            String firstHeader = getCellValueAsString(firstRow, columnIndex)
            String secondHeader = getCellValueAsString(secondRow, columnIndex)

            String namespace = secondHeader ? firstHeader ?: getCellValueAsString(firstRow, mergeRegion?.firstColumn) : null
            String key = secondHeader ?: firstHeader

            metadata += new MetadataColumn(namespace: namespace, key: key, value: getCellValueAsString(it))
        }
    }

    void addMetadataToRow(Row row) {
        if (metadata) {
            Row firstRow = row.sheet.getRow(METADATA_NAMESPACE_ROW)
            Row secondRow = row.sheet.getRow(METADATA_KEY_ROW)

            metadata.groupBy {it.namespace}.each {ns, mds ->

                Cell namespaceHeader = findCell(firstRow, ns)
                CellRangeAddress mergeRegion = getMergeRegion(namespaceHeader)

                mds.each {md ->
                    logger.debug('Adding {}:{}', md.namespace, md.key)

                    Cell keyHeader

                    if (mergeRegion) {
                        keyHeader = findCell(secondRow, md.key, mergeRegion.firstColumn, mergeRegion.lastColumn)
                    } else {
                        keyHeader = findCell(secondRow, md.key, namespaceHeader.columnIndex)
                    }

                    addCellToRow(row, keyHeader.columnIndex, md.value, true)
                }
            }
        }
    }

    void addToMetadata(String key, String value) {
        metadata += new MetadataColumn(key: key, value: value)
    }

    static Map<String, Set<String>> getMetadataNamespaceAndKeys(List<DataRow> dataRows) {
        dataRows.collect {it.metadata}
                .flatten()
                .groupBy {it.namespace}
                .collectEntries {[it.key, it.value.key as Set]}.findAll {it.value} as Map<String, Set<String>>
    }
}
