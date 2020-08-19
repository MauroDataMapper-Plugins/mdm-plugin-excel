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

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElement
import org.apache.poi.ss.usermodel.Row

import static uk.ac.ox.softeng.maurodatamapper.plugins.excel.ExcelPlugin.DATACLASS_PATH_SPLIT_REGEX

/**
 * @since 09/03/2018
 */
class ElementDataRow extends uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.EnumerationDataRow {

    String sourceDataClassPath
    String sourceDataElementName
    String targetDataElementName
    String targetDataClassPath

    ElementDataRow() {
        super()
    }

    ElementDataRow(DataElement source, DataElement target) {
        this()
        if (source) {
            sourceDataClassPath = buildPath(source.dataClass)
            sourceDataElementName = source.label
        }
        if (target) {
            targetDataClassPath = buildPath(target.dataClass)
            targetDataElementName = target.label
        }
    }

    @Override
    void setAndInitialise(Row row) {
        setRow(row)
        sourceDataClassPath = getCellValueAsString(row, 0)
        sourceDataElementName = getCellValueAsString(row, 1)

        targetDataElementName = getCellValueAsString(row, 4)
        targetDataClassPath = getCellValueAsString(row, 5)
    }

    boolean hasSource() {
        sourceDataElementName
    }

    boolean hasTarget() {
        targetDataElementName
    }

    List<String> getSourceDataClassPathList() {
        sourceDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    List<String> getTargetDataClassPathList() {
        targetDataClassPath?.split(DATACLASS_PATH_SPLIT_REGEX)?.toList()*.trim()
    }

    String buildPath(DataClass dataClass) {
        if (!dataClass.parentDataClass) return dataClass.label
        "${buildPath(dataClass.parentDataClass)}|${dataClass.label}"
    }

    @Override
    Row buildRow(Row row) {
        addCellToRow(row, 0, sourceDataClassPath)
        addCellToRow(row, 1, sourceDataElementName)
        addCellToRow(row, 4, targetDataElementName)
        addCellToRow(row, 5, targetDataClassPath)
        row
    }
}
