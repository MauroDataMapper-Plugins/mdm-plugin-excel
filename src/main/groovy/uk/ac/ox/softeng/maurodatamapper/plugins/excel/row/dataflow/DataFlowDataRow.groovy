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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow

import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlow

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column.MetadataColumn

import org.apache.poi.ss.usermodel.Row

/**
 * @since 09/03/2018
 */
class DataFlowDataRow extends uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.StandardDataRow {

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
