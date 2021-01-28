/*
 * Copyright 2020 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.slf4j.Logger
import org.slf4j.LoggerFactory

@Slf4j
@CompileStatic
trait CellHandler {

    final Logger log = LoggerFactory.getLogger(getClass())

    private final DataFormatter dataFormatter = new DataFormatter()

    String getCellValue(Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    String getCellValue(Row row, Integer column) {
        row && column != null ? getCellValue(row.getCell(column)) : ''
    }

    CellRangeAddress getMergeRegion(Cell cell) {
        getMergeRegion(cell?.sheet, cell.rowIndex, cell.columnIndex)
    }

    CellRangeAddress getMergeRegion(Sheet sheet, int rowIndex, int columnIndex) {
        sheet?.mergedRegions?.find { it.isInRange(rowIndex, columnIndex) }
    }
}
