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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.workbook

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.ExcelPlugin
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.DataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.EnumerationValueDataRow

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide

import groovy.util.logging.Slf4j

@Slf4j
// @CompileStatic
trait WorkbookExporter extends WorkbookHandler {

    final String fileExtension = 'xlsx'
    final String fileType = 'application/vnd.ms-excel'
    final String templateContentSheet = 'KEY_1'

    XSSFCellStyle createDefaultCellStyle(XSSFWorkbook workbook) {
        XSSFColor borderColour = new XSSFColor(ExcelPlugin.BORDER_COLOUR).tap {
            setTint ExcelPlugin.BORDER_COLOUR_TINT
        }
        workbook.createCellStyle().tap {
            setBorderTop BorderStyle.THIN
            setBorderLeft BorderStyle.THIN
            setBorderBottom BorderStyle.THIN
            setBorderRight BorderStyle.THIN
            setBorderColor BorderSide.TOP, borderColour
            setBorderColor BorderSide.LEFT, borderColour
            setBorderColor BorderSide.BOTTOM, borderColour
            setBorderColor BorderSide.RIGHT, borderColour
        }
    }

    void addMetadataHeadersToSheet(Sheet sheet, List<DataRow> dataRows, int cellStyleColumn = 0) {
        Map<String, Set<String>> metadataMapping = DataRow.getMetadataNamespaceAndKeys(dataRows)
        Row firstRow = sheet.getRow(0)
        Row secondRow = sheet.getRow(1)
        int namespaceColumnIndex = firstRow.lastCellNum

        metadataMapping.sort().each { String namespace, Set<String> keys ->
            Cell namespaceHeader = buildHeaderCell(firstRow, namespaceColumnIndex, cellStyleColumn, namespace, false)
            if (keys.size() == 1) {
                buildHeaderCell(secondRow, namespaceHeader.columnIndex, cellStyleColumn, keys.first(), false)
                namespaceColumnIndex++
            } else {
                sheet.addMergedRegion(new CellRangeAddress(
                    DataRow.METADATA_NAMESPACE_ROW, DataRow.METADATA_NAMESPACE_ROW,
                    namespaceHeader.columnIndex, namespaceHeader.columnIndex + keys.size() - 1))
                keys.eachWithIndex { String key, int i ->
                    buildHeaderCell(secondRow, namespaceHeader.columnIndex + i, cellStyleColumn, key, false)
                }
                namespaceColumnIndex += keys.size()
            }
        }
    }

    Sheet createContentSheetFromTemplate(XSSFWorkbook workbook, String name) {
        try {
            workbook.cloneSheet(workbook.getSheetIndex(templateContentSheet), name)
        } catch (IllegalArgumentException ignored) {
            createContentSheetFromTemplate(workbook, "${name}.1")
        }
    }

    Sheet configureSheetStyle(Sheet sheet, int metadataColumnIndex, int numberOfHeaderRows, XSSFCellStyle defaultStyle, XSSFCellStyle colouredStyle) {
        XSSFCellStyle metadataBorderStyle = sheet.workbook.createCellStyle().tap {
            cloneStyleFrom defaultStyle
            setBorderRight BorderStyle.DOUBLE
        } as XSSFCellStyle
        XSSFCellStyle metadataColouredBorderStyle = sheet.workbook.createCellStyle().tap {
            cloneStyleFrom colouredStyle
            setBorderRight BorderStyle.DOUBLE
        } as XSSFCellStyle

        int lastColumnIndex = sheet.max { it.lastCellNum }.lastCellNum
        boolean colourRow = false

        for (int rowNumber = numberOfHeaderRows; rowNumber <= sheet.lastRowNum; rowNumber++) {
            Row row = sheet.getRow(rowNumber)
            if (!row) continue

            CellStyle mainCellStyle = colourRow ? colouredStyle : defaultStyle
            CellStyle borderCellStyle = colourRow ? metadataColouredBorderStyle : metadataBorderStyle

            setRowStyle(row, mainCellStyle, borderCellStyle, 0, lastColumnIndex, metadataColumnIndex)
            rowNumber = configureExtraRowStyle(row, metadataColumnIndex, lastColumnIndex, defaultStyle, colouredStyle, mainCellStyle, borderCellStyle,
                                               metadataBorderStyle, metadataColouredBorderStyle, colourRow)
            colourRow = !colourRow
        }

        List<Integer> columnSizes = sheet.sheetName == ExcelPlugin.DATAMODELS_SHEET_NAME ? [3, 4, 5] : [5, 7, 9]
        autoSizeColumns(sheet, 0, 1, *columnSizes)
        autoSizeHeaderColumnsAfter(sheet, metadataColumnIndex)
        sheet
    }

    void loadDataRowsIntoSheet(Sheet sheet, List<DataRow> dataRows) {
        log.debug('Loading DataRows into sheet')
        addMetadataHeadersToSheet(sheet, dataRows, 3)
        dataRows.each { DataRow dataRow ->
            log.debug('Adding to row {}', sheet.lastRowNum + 1)
            dataRow.buildRow(sheet.createRow(sheet.lastRowNum + 1))
        }
    }

    XSSFWorkbook removeTemplateSheet(XSSFWorkbook workbook) {
        log.debug('Removing the template sheet')
        workbook.tap {
            removeSheetAt(workbook.getSheetIndex(templateContentSheet))
            setActiveSheet(0)
        }
    }

    private void setRowStyle(Row row, XSSFCellStyle cellStyle, XSSFCellStyle borderCellStyle, int start, int finish, int metadataColumnIndex) {
        CellStyle workbookDefaultStyle = row.sheet.workbook.getCellStyleAt(0)
        (start..<finish).each { int cellNumber ->
            Cell cell = row.getCell(cellNumber)
            CellStyle styleToUse = cellNumber == metadataColumnIndex - 1 ? borderCellStyle : cellStyle
            if (!cell) {
                cell = CellUtil.getCell(row, cellNumber)
                cell.setCellStyle(styleToUse)
            } else if (!cell.cellStyle || cell.cellStyle == workbookDefaultStyle) {
                cell.setCellStyle(styleToUse)
            } else if (cell.cellStyle.wrapText) {
                copyCellStyleInto(cell, styleToUse)
            }
        }
    }

    private void copyCellStyleInto(Cell cell, XSSFCellStyle style) {
        if (style.fillPatternEnum == FillPatternType.SOLID_FOREGROUND) {
            CellUtil.setCellStyleProperty(cell, CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND)
            (cell.cellStyle as XSSFCellStyle).setFillForegroundColor(style.fillForegroundColorColor)
        }
        XSSFColor borderColour = style.leftBorderXSSFColor
        (cell.cellStyle as XSSFCellStyle).tap {
            setBorderTop BorderStyle.THIN
            setBorderLeft BorderStyle.THIN
            setBorderBottom BorderStyle.THIN
            setBorderRight BorderStyle.THIN
            setBorderColor BorderSide.TOP, borderColour
            setBorderColor BorderSide.LEFT, borderColour
            setBorderColor BorderSide.BOTTOM, borderColour
            setBorderColor BorderSide.RIGHT, borderColour
        }
    }

    private Integer configureExtraRowStyle(Row row, int metadataColumnIndex, int lastColumnIndex, XSSFCellStyle defaultStyle,
                                           XSSFCellStyle colouredStyle, XSSFCellStyle mainCellStyle, XSSFCellStyle borderCellStyle,
                                           XSSFCellStyle metadataBorderStyle, XSSFCellStyle metadataColouredBorderStyle, boolean colourRow) {
        CellRangeAddress mergedRegion = getMergeRegion(row.getCell(0))
        if (!mergedRegion) return row.rowNum

        (mergedRegion.firstRow..mergedRegion.lastRow).each { int rowNumber ->
            Row sheetRow = row.sheet.getRow(rowNumber)
            setRowStyle(sheetRow, mainCellStyle, borderCellStyle, 0, EnumerationValueDataRow.KEY_COL_INDEX, metadataColumnIndex)
            setRowStyle(sheetRow, colourRow ? colouredStyle : defaultStyle, colourRow ? metadataColouredBorderStyle : metadataBorderStyle,
                        EnumerationValueDataRow.KEY_COL_INDEX, EnumerationValueDataRow.VALUE_COL_INDEX + 1, metadataColumnIndex)
            setRowStyle(sheetRow, mainCellStyle, borderCellStyle, EnumerationValueDataRow.VALUE_COL_INDEX + 1, lastColumnIndex, metadataColumnIndex)
            colourRow = !colourRow
        }
        mergedRegion.lastRow
    }

    private void autoSizeColumns(Sheet sheet, Integer... columns) {
        columns.each { sheet.autoSizeColumn(it, true) }
    }

    private void autoSizeHeaderColumnsAfter(Sheet sheet, int columnIndex) {
        sheet.getRow(0)
             .findAll { it.columnIndex >= columnIndex && getCellValue(it) }
             .each { sheet.autoSizeColumn(it.columnIndex, true) }
    }
}
