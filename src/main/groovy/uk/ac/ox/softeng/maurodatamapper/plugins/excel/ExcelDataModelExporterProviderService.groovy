/*
 * Copyright 2020-2021 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.beans.factory.annotation.Autowired

@Slf4j
@CompileStatic
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
        getClass().getPackage().getSpecificationVersion() ?: 'SNAPSHOT'
    }

    @Override
    String getFileType() {
        ExcelPlugin.EXCEL_FILETYPE
    }

    @Override
    String getFileExtension() {
        ExcelPlugin.EXCEL_FILE_EXTENSION
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
        workbook = loadWorkbookFromFilename(ExcelPlugin.DATAMODELS_IMPORT_TEMPLATE_FILENAME) as XSSFWorkbook
        loadDataModelsIntoWorkbook(dataModels).withCloseable { XSSFWorkbook workbook ->
            if (!workbook) return null
            new ByteArrayOutputStream().tap { ByteArrayOutputStream exportStream ->
                workbook.write(exportStream)
                log.info('DataModels exported')
            }
        }
    }

    private XSSFWorkbook loadDataModelsIntoWorkbook(List<DataModel> dataModels) {
        loadWorkbookCellAndBorderStyles()
        List<DataModelDataRow> dataRows = dataModels.collect { new DataModelDataRow(it) }
        Sheet dataModelSheet = workbook.getSheet(ExcelPlugin.DATAMODELS_SHEET_NAME).tap { Sheet sheet ->
            addMetadataHeadersToSheet sheet, dataRows
        }
        dataRows.each { DataModelDataRow dataRow ->
            loadDataModelDataRowToSheet dataModelSheet, dataRow
            loadContentDataRowsToSheet dataRow
        }
        removeTemplateSheet()
    }

    private void loadDataModelDataRowToSheet(Sheet sheet, DataModelDataRow dataRow) {
        log.debug('Adding row to DataModel [{}] sheet', sheet.sheetName)
        dataRow.buildRow(sheet.createRow(sheet.lastRowNum + 1))
        configureDataModelSheetStyle(sheet, dataRow)
    }

    private void loadContentDataRowsToSheet(DataModelDataRow dataRow) {
        log.debug('Adding content rows to DataModel [{}] sheet', dataRow.sheetKey)
        List<ContentDataRow> contentDataRows = createContentDataRows(dataRow.dataModel.id, dataRow.dataModel.childDataClasses.sort {it.label})
        if (!contentDataRows) return
        loadContentSheet(dataRow).with { Sheet contentSheet ->
            loadDataRowsIntoSheet contentSheet, contentDataRows
            configureContentSheetStyle contentSheet, contentDataRows
        }
    }

    private Sheet loadContentSheet(DataModelDataRow dataRow) {
        log.debug('Loading DataModel [{}] content sheet', dataRow.name)
        createContentSheetFromTemplate(dataRow.sheetKey).tap {Sheet contentSheet ->
            dataRow.sheetKey = contentSheet.sheetName
        }
    }

    private <K extends CatalogueItem> List<ContentDataRow> createContentDataRows(UUID dataModelId, List<K> catalogueItems,
                                                                                 List<ContentDataRow> dataRows = []) {
        log.debug('Creating content rows')
        catalogueItems.each {CatalogueItem catalogueItem ->
            getLog().trace('Creating content {} : {}', catalogueItem.domainType, catalogueItem.label)
            dataRows << new ContentDataRow(catalogueItem)
            if (catalogueItem instanceof DataClass) {
                dataRows = createContentDataRows(dataModelId, dataElementService.findAllByDataClassIdJoinDataType(catalogueItem.id), dataRows)
                dataRows = createContentDataRows(dataModelId,
                                                 dataClassService.findAllByDataModelIdAndParentDataClassId(dataModelId, catalogueItem.id,
                                                                                                           [sort: 'label']) as List<DataClass>,
                                                 dataRows)
            }
        }
        dataRows
    }

    private void configureDataModelSheetStyle(Sheet sheet, DataModelDataRow dataRow) {
        log.debug('Configuring DataModel [{}] sheet style', sheet.sheetName)
        configureSheetStyle(sheet, dataRow.firstMetadataColumn, ExcelPlugin.DATAMODELS_NUM_HEADER_ROWS)
    }

    private void configureContentSheetStyle(Sheet sheet, List<ContentDataRow> dataRows) {
        log.debug('Configuring DataModel [{}] content sheet style', sheet.sheetName)
        configureSheetStyle(sheet, dataRows.first().firstMetadataColumn, ExcelPlugin.CONTENT_NUM_HEADER_ROWS)
    }
}
