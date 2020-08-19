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

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationValue
import org.apache.poi.ss.usermodel.Row

/**
 * @since 01/03/2018
 */
class EnumerationValueDataRow extends uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.EnumerationDataRow {

    public static final int KEY_COL_INDEX = 8
    public static final int VALUE_COL_INDEX = 9

    EnumerationValueDataRow() {
        super()
    }

    EnumerationValueDataRow(EnumerationValue enumerationValue) {
        this()
        key = enumerationValue.key
        value = enumerationValue.value
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)
        key = getCellValueAsString(row, KEY_COL_INDEX)
        value = getCellValueAsString(row, VALUE_COL_INDEX)
    }

    Row buildRow(Row row) {
        addCellToRow(row, KEY_COL_INDEX, key)
        addCellToRow(row, VALUE_COL_INDEX, value)
        row
    }
}
