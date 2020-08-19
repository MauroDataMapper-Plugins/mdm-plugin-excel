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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.util

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.DataRow

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
import org.slf4j.Logger

import static uk.ac.ox.softeng.maurodatamapper.plugins.excel.ExcelPlugin.BORDER_COLOUR
import static uk.ac.ox.softeng.maurodatamapper.plugins.excel.ExcelPlugin.BORDER_COLOUR_TINT
import static uk.ac.ox.softeng.maurodatamapper.plugins.excel.ExcelPlugin.DATAMODELS_SHEET_NAME
import static uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.DataRow.getMetadataNamespaceAndKeys

/**
 * @since 13/03/2018
 */
trait WorkbookExporter implements WorkbookHandler {

    abstract Logger getLogger()

    abstract Integer configureExtraRowStyle(Row row, int metadataColumnIndex, int lastColumnIndex, XSSFCellStyle defaultStyle,
                                            XSSFCellStyle colouredStyle, XSSFCellStyle mainCellStyle, XSSFCellStyle borderCellStyle,
                                            XSSFCellStyle metadataBorderStyle, XSSFCellStyle metadataColouredBorderStyle, boolean colourRow)

    String getFileExtension() {
        'xlsx'
    }

    String getFileType() {
        'application/vnd.ms-excel'
    }

    String getTemplateContentSheet() {
        'KEY_1'
    }

    Sheet createComponentSheetFromTemplate(XSSFWorkbook workbook, String name) {
        try {
            workbook.cloneSheet(workbook.getSheetIndex(templateContentSheet), name)
        } catch (IllegalArgumentException ignored) {
            createComponentSheetFromTemplate(workbook, "${name}.1")
        }
    }

    XSSFWorkbook removeTemplateSheet(XSSFWorkbook workbook) {
        logger.debug('Removing the template sheet')
        workbook.removeSheetAt(workbook.getSheetIndex(templateContentSheet))
        workbook.setActiveSheet(0)
        workbook
    }

    void loadDataRowsIntoSheet(Sheet sheet, List<DataRow> dataRows) {
        logger.debug('Loading DataRows into sheet')

        addMetadataHeadersToSheet(sheet, dataRows, 3)

        dataRows.each {dataRow ->
            logger.debug('Adding to row {}', sheet.lastRowNum + 1)
            dataRow.buildRow(sheet.createRow(sheet.lastRowNum + 1))
        }
    }

    void addMetadataHeadersToSheet(Sheet sheet, List<DataRow> dataRows, int cellStyleColumn = 0) {
        Map<String, Set<String>> metadataMapping = getMetadataNamespaceAndKeys(dataRows)

        Row firstRow = sheet.getRow(0)
        Row secondRow = sheet.getRow(1)

        int nsColumnIndex = firstRow.lastCellNum

        metadataMapping.sort().each {namespace, keys ->

            Cell namespaceHeader = buildHeaderCell(firstRow, nsColumnIndex, cellStyleColumn, namespace, false)

            if (keys.size() == 1) {
                buildHeaderCell(secondRow, namespaceHeader.columnIndex, cellStyleColumn, keys.first(), false)
                nsColumnIndex++
            } else {

                CellRangeAddress mergeRegion = new CellRangeAddress(DataRow.METADATA_NAMESPACE_ROW, DataRow.METADATA_NAMESPACE_ROW,
                                                                    namespaceHeader.columnIndex,
                                                                    namespaceHeader.columnIndex + keys.size() - 1)
                sheet.addMergedRegion(mergeRegion)

                keys.eachWithIndex {key, i ->
                    buildHeaderCell(secondRow, namespaceHeader.columnIndex + i, cellStyleColumn, key, false)
                }
                nsColumnIndex += keys.size()
            }
        }
    }

    Sheet configureSheetStyle(Sheet sheet, int metadataColumnIndex, int numberOfHeaderRows, XSSFCellStyle defaultStyle,
                              XSSFCellStyle colouredStyle) {

        XSSFCellStyle metadataBorderStyle = sheet.workbook.createCellStyle() as XSSFCellStyle
        metadataBorderStyle.cloneStyleFrom(defaultStyle)
        metadataBorderStyle.setBorderRight(BorderStyle.DOUBLE)

        XSSFCellStyle metadataColouredBorderStyle = sheet.workbook.createCellStyle() as XSSFCellStyle
        metadataColouredBorderStyle.cloneStyleFrom(colouredStyle)
        metadataColouredBorderStyle.setBorderRight(BorderStyle.DOUBLE)

        int lastColumnIndex = sheet.max {it.lastCellNum}.lastCellNum
        boolean colourRow = false

        for (int r = numberOfHeaderRows; r <= sheet.lastRowNum; r++) {
            Row row = sheet.getRow(r)

            if (!row) continue

            CellStyle mainCellStyle = colourRow ? colouredStyle : defaultStyle
            CellStyle borderCellStyle = colourRow ? metadataColouredBorderStyle : metadataBorderStyle

            setRowStyle(row, mainCellStyle, borderCellStyle, 0, lastColumnIndex, metadataColumnIndex)

            r = configureExtraRowStyle(row, metadataColumnIndex, lastColumnIndex, defaultStyle,
                                       colouredStyle, mainCellStyle, borderCellStyle,
                                       metadataBorderStyle, metadataColouredBorderStyle, colourRow)

            colourRow = !colourRow
        }

        if (sheet.sheetName == DATAMODELS_SHEET_NAME) autoSizeColumns(sheet, 0, 1, 3, 4, 5)
        else autoSizeColumns(sheet, 0, 1, 5, 7, 9)

        autoSizeHeaderColumnsAfter(sheet, metadataColumnIndex)

        sheet
    }

    void setRowStyle(Row row, XSSFCellStyle cellStyle, XSSFCellStyle borderCellStyle, int start, int finish, int metadataColumnIndex) {

        CellStyle workbookDefaultStyle = row.sheet.workbook.getCellStyleAt(0)

        for (int i = start; i < finish; i++) {
            Cell cell = row.getCell(i)

            CellStyle styleToUse = i == metadataColumnIndex - 1 ? borderCellStyle : cellStyle

            if (!cell) {
                cell = CellUtil.getCell(row, i)
                cell.setCellStyle(styleToUse)
            } else if (cell.cellStyle.wrapText) {
                copyCellStyleInto(cell, styleToUse)
            } else if (!cell.cellStyle || cell.cellStyle == workbookDefaultStyle) cell.setCellStyle(styleToUse)
        }
    }

    XSSFCellStyle createDefaultCellStyle(XSSFWorkbook workbook) {
        XSSFCellStyle cellStyle = workbook.createCellStyle()
        XSSFColor borderColour = new XSSFColor(BORDER_COLOUR)
        borderColour.setTint(BORDER_COLOUR_TINT)

        cellStyle.with {
            setBorderRight BorderStyle.THIN
            setBorderLeft(BorderStyle.THIN)
            setBorderBottom(BorderStyle.THIN)
            setBorderTop(BorderStyle.THIN)
            setBorderColor(BorderSide.LEFT, borderColour)
            setBorderColor(BorderSide.TOP, borderColour)
            setBorderColor(BorderSide.RIGHT, borderColour)
            setBorderColor(BorderSide.BOTTOM, borderColour)
            it
        }
    }

    void copyCellStyleInto(Cell cell, XSSFCellStyle style) {

        if (style.fillPatternEnum == FillPatternType.SOLID_FOREGROUND) {
            CellUtil.setCellStyleProperty(cell, CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND)
            (cell.cellStyle as XSSFCellStyle).setFillForegroundColor(style.fillForegroundColorColor)
        }
        XSSFColor borderColour = style.leftBorderXSSFColor
        (cell.cellStyle as XSSFCellStyle).with {
            setBorderRight BorderStyle.THIN
            setBorderLeft(BorderStyle.THIN)
            setBorderBottom(BorderStyle.THIN)
            setBorderTop(BorderStyle.THIN)
            setBorderColor(BorderSide.LEFT, borderColour)
            setBorderColor(BorderSide.TOP, borderColour)
            setBorderColor(BorderSide.RIGHT, borderColour)
            setBorderColor(BorderSide.BOTTOM, borderColour)
            it
        }
    }
}
