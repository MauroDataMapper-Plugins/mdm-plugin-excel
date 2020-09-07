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

import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.StandardDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column.MetadataColumn

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
                if (word.length() > 0) {
                    sheetKey += word[0].toUpperCase()
                }
            }
        }
        name = dataModel.label
        description = dataModel.description
        author = dataModel.author
        organisation = dataModel.organisation
        type = dataModel.modelType

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
