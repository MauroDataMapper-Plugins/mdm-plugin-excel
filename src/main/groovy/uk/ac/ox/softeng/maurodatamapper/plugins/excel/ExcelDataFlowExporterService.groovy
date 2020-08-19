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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel

import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlow
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowComponent
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowService
import ox.softeng.metadatacatalogue.core.spi.exporter.DataFlowExporterPlugin
import ox.softeng.metadatacatalogue.core.user.CatalogueUser

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.util.WorkbookExporter

import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.beans.factory.annotation.Autowired

import static ExcelPlugin.ALTERNATING_COLUMN_COLOUR
import static ExcelPlugin.CELL_COLOUR_TINT
import static ExcelPlugin.DATAFLOWCOMPONENT_HEADER_ROWS
import static ExcelPlugin.DATAFLOWS_HEADER_ROWS
import static ExcelPlugin.DATAFLOWS_SHEET_NAME
import static ExcelPlugin.DATAFLOW_TEMPLATE_FILENAME

/**
 * @since 01/03/2018
 */
class ExcelDataFlowExporterService implements DataFlowExporterPlugin, WorkbookExporter {

    @Autowired
    DataFlowService dataFlowService

    @Override
    String getDisplayName() {
        'Excel (XLSX) DataFlow Exporter'
    }

    @Override
    String getVersion() {
        '1.0.0'
    }

    @Override
    ByteArrayOutputStream exportDataFlow(CatalogueUser currentUser, DataFlow dataFlow) throws ApiException {
        exportDataFlows(currentUser, [dataFlow])
    }

    @Override
    ByteArrayOutputStream exportDataFlows(CatalogueUser currentUser, List<DataFlow> dataFlows) throws ApiException {

        logger.info('Exporting DataFlows to Excel')
        XSSFWorkbook workbook = null
        try {
            workbook = loadWorkbookFromFilename(DATAFLOW_TEMPLATE_FILENAME) as XSSFWorkbook

            workbook = loadDataFlowDataRowsIntoWorkbook(dataFlows.collect {new uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow
                .DataFlowDataRow(it)}, workbook)

            if (workbook) {
                ByteArrayOutputStream os = new ByteArrayOutputStream()
                workbook.write(os)
                logger.info('DataFlows exported')
                return os
            }
        } finally {
            closeWorkbook(workbook)
        }
        null
    }

    XSSFWorkbook loadDataFlowDataRowsIntoWorkbook(List<uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowDataRow> dataFlowDataRows, XSSFWorkbook workbook) {
        XSSFColor cellColour = new XSSFColor(ALTERNATING_COLUMN_COLOUR)
        cellColour.setTint(CELL_COLOUR_TINT)

        XSSFCellStyle defaultStyle = createDefaultCellStyle(workbook as XSSFWorkbook)

        XSSFCellStyle colouredStyle = workbook.createCellStyle() as XSSFCellStyle
        colouredStyle.cloneStyleFrom(defaultStyle)
        colouredStyle.setFillForegroundColor(cellColour)
        colouredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

        Sheet dataFlowSheet = workbook.getSheet(DATAFLOWS_SHEET_NAME)

        addMetadataHeadersToSheet(dataFlowSheet, dataFlowDataRows)

        dataFlowDataRows.each {dataRow ->
            logger.debug('Loading DataFlow {} into workbook', dataRow.name)

            // We need to allow for potential sheet name already exists
            Sheet componentSheet = createComponentSheetFromTemplate(workbook, dataRow.sheetKey)
            dataRow.sheetKey = componentSheet.sheetName

            logger.debug('Adding DataFlow row')
            dataRow.buildRow(dataFlowSheet.createRow(dataFlowSheet.lastRowNum + 1))

            logger.debug('Configuring {} sheet style', dataFlowSheet.sheetName)
            configureSheetStyle(dataFlowSheet, dataRow.firstMetadataColumn, DATAFLOWS_HEADER_ROWS, defaultStyle, colouredStyle)

            logger.debug('Adding DataFlow content to sheet {}', dataRow.sheetKey)
            List<uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowComponentDataRow> componentDataRows = createComponentDataRows(dataRow.dataFlow.dataFlowComponents.sort {it.label})

            if (componentDataRows) {

                loadDataRowsIntoSheet(componentSheet, componentDataRows)

                logger.debug('Configuring {} sheet style', componentSheet.sheetName)
                configureSheetStyle(componentSheet, componentDataRows.first().firstMetadataColumn, DATAFLOWCOMPONENT_HEADER_ROWS, defaultStyle,
                                    colouredStyle)
            } else {
                logger.warn('No component rows to add for DataFlow')
            }
        }
        removeTemplateSheet(workbook)
    }

    List<uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowComponentDataRow> createComponentDataRows(List<DataFlowComponent> dataFlowComponents,
                                                                                                                       List<uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowComponentDataRow> componentDataRows = []) {
        logger.debug('Creating content DataRows')
        dataFlowComponents.each {dfc ->
            logger.trace('Creating content {}:{}', dfc.domainType, dfc.label)
            componentDataRows += new uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowComponentDataRow(dfc)
        }
        componentDataRows
    }

    @Override
    Integer configureExtraRowStyle(Row row, int metadataColumnIndex, int lastColumnIndex, XSSFCellStyle defaultStyle, XSSFCellStyle colouredStyle,
                                   XSSFCellStyle mainCellStyle, XSSFCellStyle borderCellStyle, XSSFCellStyle metadataBorderStyle,
                                   XSSFCellStyle metadataColouredBorderStyle, boolean colourRow) {
        CellRangeAddress mergedRegion = getMergeRegion(row.getCell(2))
        if (mergedRegion) {
            boolean internalColourRow = colourRow
            for (int m = mergedRegion.firstRow; m <= mergedRegion.lastRow; m++) {
                Row internalRow = row.sheet.getRow(m)

                setRowStyle(internalRow, internalColourRow ? colouredStyle : defaultStyle,
                            internalColourRow ? metadataColouredBorderStyle : metadataBorderStyle,
                            0, 2, metadataColumnIndex)

                setRowStyle(internalRow, mainCellStyle, borderCellStyle, 2, 4, metadataColumnIndex)

                setRowStyle(internalRow, internalColourRow ? colouredStyle : defaultStyle,
                            internalColourRow ? metadataColouredBorderStyle : metadataBorderStyle,
                            4, 6, metadataColumnIndex)

                setRowStyle(internalRow, mainCellStyle, borderCellStyle, 6, lastColumnIndex, metadataColumnIndex)

                internalColourRow = !internalColourRow
            }
            return mergedRegion.lastRow
        }
        row.rowNum
    }
}
