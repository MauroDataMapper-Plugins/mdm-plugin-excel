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

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClassService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElementService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.ContentDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.DataModelDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.workbook.WorkbookExporter
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.beans.factory.annotation.Autowired

import groovy.util.logging.Slf4j

@Slf4j
// @CompileStatic
class ExcelDataModelExporterProviderService extends DataModelExporterProviderService implements WorkbookExporter {

    @Autowired
    DataModelService dataModelService

    @Autowired
    DataClassService dataClassService

    @Autowired
    DataElementService dataElementService

    @Override
    String getDisplayName() {
        'Excel (XLSX) Exporter'
    }

    @Override
    String getVersion() {
        '2.0.0-SNAPSHOT'
    }

    @Override
    Boolean canExportMultipleDomains() {
        true
    }

    @Override
    ByteArrayOutputStream exportDataModel(User currentUser, DataModel dataModel) throws ApiException {
        exportDataModels(currentUser, [dataModel])
    }

    @Override
    ByteArrayOutputStream exportDataModels(User currentUser, List<DataModel> dataModels) throws ApiException {
        log.info('Exporting DataModels to Excel')
        XSSFWorkbook workbook = null
        try {
            workbook = loadDataModelDataRowsIntoWorkbook(
                dataModels.collect { new DataModelDataRow(it) },
                loadWorkbookFromFilename(ExcelPlugin.DATAMODEL_TEMPLATE_FILENAME) as XSSFWorkbook)
            if (!workbook) return null

            ByteArrayOutputStream exportStream = new ByteArrayOutputStream()
            workbook.write(exportStream)
            log.info('DataModels exported')
            exportStream
        } finally {
            closeWorkbook(workbook)
        }
    }

    private XSSFWorkbook loadDataModelDataRowsIntoWorkbook(List<DataModelDataRow> dataRows, XSSFWorkbook workbook) {
        XSSFCellStyle defaultStyle = createDefaultCellStyle(workbook as XSSFWorkbook)
        XSSFColor cellColour = new XSSFColor(ExcelPlugin.ALTERNATING_COLUMN_COLOUR).tap {
            tint = ExcelPlugin.CELL_COLOUR_TINT
        }
        XSSFCellStyle colouredStyle = workbook.createCellStyle().tap {
            cloneStyleFrom(defaultStyle)
            fillForegroundColor = cellColour
            fillPattern = FillPatternType.SOLID_FOREGROUND
        } as XSSFCellStyle

        Sheet dataModelSheet = workbook.getSheet(ExcelPlugin.DATAMODELS_SHEET_NAME)
        addMetadataHeadersToSheet(dataModelSheet, dataRows)

        dataRows.each { DataModelDataRow dataRow ->
            log.debug('Loading DataModel {} into workbook', dataRow.name)
            Sheet contentSheet = createComponentSheetFromTemplate(workbook, dataRow.sheetKey)
            dataRow.sheetKey = contentSheet.sheetName

            log.debug('Adding DataModel row')
            dataRow.buildRow(dataModelSheet.createRow(dataModelSheet.lastRowNum + 1))

            log.debug('Configuring {} sheet style', dataModelSheet.sheetName)
            configureSheetStyle(dataModelSheet, dataRow.firstMetadataColumn, ExcelPlugin.DATAMODELS_HEADER_ROWS, defaultStyle, colouredStyle)

            log.debug('Adding DataModel content to sheet {}', dataRow.sheetKey)
            List<ContentDataRow> contentDataRows = createContentDataRows(dataRow.dataModel.childDataClasses.sort { it.label })
            if (!contentDataRows) log.warn('No content rows to add for DataModel')
            else {
                loadDataRowsIntoSheet(contentSheet, contentDataRows)
                log.debug('Configuring {} sheet style', contentSheet.sheetName)
                configureSheetStyle(contentSheet, contentDataRows.first().firstMetadataColumn, ExcelPlugin.CONTENT_HEADER_ROWS, defaultStyle,
                                    colouredStyle)
            }
        }
        removeTemplateSheet(workbook)
    }

    private List<ContentDataRow> createContentDataRows(List<CatalogueItem> catalogueItems, List<ContentDataRow> dataRows = []) {
        log.debug('Creating content DataRows')
        catalogueItems.each { CatalogueItem catalogueItem ->
            log.trace('Creating content {}:{}', catalogueItem.domainType, catalogueItem.label)
            dataRows << new ContentDataRow(catalogueItem)
            if (catalogueItem instanceof DataClass) {
                dataRows = createContentDataRows(dataElementService.findAllByDataClassIdJoinDataType(catalogueItem.id), dataRows)
                dataRows = createContentDataRows(dataClassService.findAllByParentDataClassId(catalogueItem.id, [sort: 'label']), dataRows)
            }
        }
        dataRows
    }
}
