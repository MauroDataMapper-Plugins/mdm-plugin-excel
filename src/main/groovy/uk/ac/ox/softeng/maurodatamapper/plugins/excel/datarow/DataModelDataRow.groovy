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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow

import uk.ac.ox.softeng.maurodatamapper.core.facet.Metadata
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel

import org.apache.poi.ss.usermodel.Row

class DataModelDataRow extends StandardDataRow {

    DataModel dataModel
    String sheetKey
    String name
    String description
    String author
    String organisation
    String type

    DataModelDataRow() {
    }

    DataModelDataRow(DataModel dataModel) {
        this.dataModel = dataModel
        String[] words = dataModel.label.split(' ')
        if (words.size() == 1) sheetKey = dataModel.label.toUpperCase()
        else {
            StringBuffer sheetKeyString = new StringBuffer()
            words.each { String word ->
                if (word.length()) sheetKeyString.append(word[0].toUpperCase())
            }
            sheetKey = sheetKeyString.toString()
        }
        name = dataModel.label
        description = dataModel.description
        author = dataModel.author
        organisation = dataModel.organisation
        type = dataModel.modelType
        dataModel.metadata.each { Metadata metadataEntry ->
            metadata << new MetadataColumn(namespace: metadataEntry.namespace, key: metadataEntry.key, value: metadataEntry.value)
        }
    }

    @Override
    Integer getFirstMetadataColumn() {
        6
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)
        sheetKey = getCellValue(row, 0)
        name = getCellValue(row, 1)
        description = getCellValue(row, 2)
        author = getCellValue(row, 3)
        organisation = getCellValue(row, 4)
        type = getCellValue(row, 5)
        extractMetadataFromColumnIndex()
    }

    @Override
    Row buildRow(Row row) {
        addCellToRow row, 0, sheetKey
        addCellToRow row, 1, name
        addCellToRow row, 2, description, true
        addCellToRow row, 3, author
        addCellToRow row, 4, organisation
        addCellToRow row, 5, type
        addMetadataToRow row
        row
    }
}
