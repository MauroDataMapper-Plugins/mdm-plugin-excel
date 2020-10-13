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

import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.enumeration.EnumerationValue

import org.apache.poi.ss.usermodel.Row

class EnumerationValueDataRow extends EnumerationDataRow {

    static final int KEY_COLUMN_INDEX = 8
    static final int VALUE_COLUMN_INDEX = 9

    EnumerationValueDataRow() {
    }

    EnumerationValueDataRow(EnumerationValue enumerationValue) {
        key = enumerationValue.key
        value = enumerationValue.value
    }

    @Override
    void initialiseRow(Row row) {
        setRow(row)
        key = getCellValue(row, KEY_COLUMN_INDEX)
        value = getCellValue(row, VALUE_COLUMN_INDEX)
    }

    @Override
    Row buildRow(Row row) {
        addCellToRow row, KEY_COLUMN_INDEX, key
        addCellToRow row, VALUE_COLUMN_INDEX, value
        row
    }

    @Override
    int getFirstMetadataColumn() {
        null
    }
}
