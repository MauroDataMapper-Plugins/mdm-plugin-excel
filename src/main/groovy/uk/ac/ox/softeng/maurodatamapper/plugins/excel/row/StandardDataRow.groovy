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

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationType

import org.apache.poi.ss.usermodel.Row

/**
 * @since 07/09/2017
 */
abstract class StandardDataRow<K extends EnumerationDataRow> extends DataRow {

    Class<K> enumerationExcelDataRowClass
    List<K> mergedContentRows

    StandardDataRow() {
        super()
        mergedContentRows = []
    }

    StandardDataRow(Class<K> enumerationExcelDataRowClass) {
        this()
        this.enumerationExcelDataRowClass = enumerationExcelDataRowClass
    }

    void addToMergedContentRows(Row row) {
        if (enumerationExcelDataRowClass) {
            K enumerationRow = enumerationExcelDataRowClass.newInstance()
            enumerationRow.setAndInitialise(row)
            mergedContentRows += enumerationRow
        }
    }

    boolean matchesEnumerationType(EnumerationType enumerationType) {
        if (!mergedContentRows) return false

        if (enumerationType.enumerationValues.size() != mergedContentRows.size()) return false

        // Check every enumeration value has an entry in the merged content rows
        enumerationType.enumerationValues.every {ev ->
            mergedContentRows.any {it.key == ev.key && it.value == ev.value}
        }
    }
}
