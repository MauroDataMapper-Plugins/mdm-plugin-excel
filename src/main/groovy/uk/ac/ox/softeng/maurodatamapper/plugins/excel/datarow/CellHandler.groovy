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

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

// @Slf4j
// @CompileStatic
trait CellHandler {

    final DataFormatter dataFormatter = new DataFormatter()

    @SuppressWarnings('GroovyAssignabilityCheck')
    void addCellToRow(Row row, int cellNum, Object content, boolean wrapText = false) {
        Cell cell = row.createCell(cellNum)
        cell.setCellValue(content)
        if (wrapText) {
            CellUtil.setCellStyleProperty(cell, CellUtil.WRAP_TEXT, wrapText)
            cell.cellStyle.setShrinkToFit(!wrapText)
        }
    }

    Cell findCell(Row row, String content) {
        row.find { getCellValue(it) == content } as Cell
    }

    Cell findCell(Row row, String content, int column) {
        findCell(row, content, column, column)
    }

    Cell findCell(Row row, String content, int firstColumn, int lastColumn) {
        row.find { it.columnIndex >= firstColumn && it.columnIndex <= lastColumn && getCellValue(it) == content } as Cell
    }

    String getCellValue(Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    String getCellValue(Row row, Integer column) {
        row && column != null ? getCellValue(row.getCell(column)) : ''
    }

    int getCellValueAsInt(Cell cell, int defaultValue) {
        String cellValue = getCellValue(cell)
        cellValue ? cellValue.toInteger() : defaultValue
    }

    CellRangeAddress getMergeRegion(Cell cell) {
        getMergeRegion(cell?.sheet, cell.rowIndex, cell.columnIndex)
    }

    CellRangeAddress getMergeRegion(Sheet sheet, int rowIndex, int columnIndex) {
        sheet?.mergedRegions?.find { it.isInRange(rowIndex, columnIndex) }
    }

    Cell buildHeaderCell(Row headerRow, int columnIndex, int cellStyleColumn, String cellValue, boolean createMergeRegion = true) {
        Cell headerCell = headerRow.createCell(columnIndex).tap {
            setCellValue cellValue
        }

        Cell copyCell = headerRow.getCell(cellStyleColumn)
        CellRangeAddress mergedRegion = getMergeRegion(copyCell)

        if (!createMergeRegion || !mergedRegion) {
            headerCell.setCellStyle(copyCell.cellStyle)
            return headerCell
        }

        CellRangeAddress newRegion = mergedRegion.copy().tap {
            firstColumn = headerCell.columnIndex
            lastColumn = headerCell.columnIndex
        }
        (newRegion.firstRow..newRegion.lastRow) { int rowNumber ->
            Row newRegionRow = headerRow.sheet.getRow(rowNumber)
            Cell newRegionHeaderCell = CellUtil.getCell(newRegionRow, headerCell.columnIndex)
            newRegionHeaderCell.setCellStyle(copyCell.cellStyle)
            if (rowNumber == newRegion.lastRow) {
                CellUtil.setCellStyleProperty(newRegionHeaderCell, CellUtil.BORDER_BOTTOM,
                                              newRegionRow.getCell(cellStyleColumn).cellStyle.borderBottom)
            }
        }
        headerRow.sheet.addMergedRegion(newRegion)
        headerCell
    }
}
