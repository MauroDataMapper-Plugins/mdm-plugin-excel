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

import ox.softeng.metadatacatalogue.core.api.exception.ApiBadRequestException
import ox.softeng.metadatacatalogue.core.api.exception.ApiUnauthorizedException
import ox.softeng.metadatacatalogue.core.catalogue.CatalogueItem
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlow
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowComponent
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowComponentService
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClassService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElement
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElementService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.spi.importer.DataFlowImporterPlugin
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.FileParameter
import ox.softeng.metadatacatalogue.core.user.CatalogueUser

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column.MetadataColumn
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowComponentDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.dataflow.DataFlowDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.util.WorkbookHandler

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.beans.factory.annotation.Autowired

import static ExcelPlugin.DATAFLOWCOMPONENT_HEADER_ROWS
import static ExcelPlugin.DATAFLOWCOMPONENT_ID_COLUMN
import static ExcelPlugin.DATAFLOWS_HEADER_ROWS
import static ExcelPlugin.DATAFLOWS_ID_COLUMN
import static ExcelPlugin.DATAFLOWS_SHEET_NAME

/**
 * @since 09/03/2018
 */
class ExcelDataFlowImporterService implements DataFlowImporterPlugin<ExcelDataFlowImporterParameters>, WorkbookHandler {

    @Autowired
    DataModelService dataModelService

    @Autowired
    DataClassService dataClassService

    @Autowired
    DataElementService dataElementService

    @Autowired
    DataFlowService dataFlowService

    @Autowired
    DataFlowComponentService dataFlowComponentService

    @Override
    String getDisplayName() {
        'Excel (XLSX) DataFlow Importer'
    }

    @Override
    String getVersion() {
        '1.0'
    }

    @Override
    Boolean canImportMultipleDomains() {
        true
    }

    @Override
    DataFlow importDataFlow(CatalogueUser currentUser, ExcelDataFlowImporterParameters params) {
        importDataFlows(currentUser, params)?.first()
    }

    @Override
    List<DataFlow> importDataFlows(CatalogueUser currentUser, ExcelDataFlowImporterParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EDFISP01', 'User must be logged in to import model')
        if (importerParameters.importFile.fileContents.size() == 0) throw new ApiBadRequestException('EDFIS02', 'Cannot import empty file')
        logger.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)

        FileParameter importFile = importerParameters.importFile

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream)

            List<DataFlowDataRow> dataFlowDataRows = loadDataRows(workbook, DataFlowDataRow, importFile.fileName,
                                                                                   DATAFLOWS_SHEET_NAME, DATAFLOWS_HEADER_ROWS,
                                                                                   DATAFLOWS_ID_COLUMN)
            if (!dataFlowDataRows) return []

            List<DataFlow> dataFlows = loadDataFlows(dataFlowDataRows, currentUser, workbook, importFile.fileName)

            return dataFlows.findAll {it.dataFlowComponents}
        } finally {
            closeWorkbook workbook
        }
    }

    List<DataFlow> loadDataFlows(List<DataFlowDataRow> dataRows, CatalogueUser catalogueUser, Workbook workbook, String filename) {
        List<DataFlow> dataFlows = []

        dataRows.each {dataRow ->
            logger.debug('Creating DataFlow {}', dataRow.name)
            DataFlow dataFlow = dataFlowService.
                replaceAndCreateDataFlow(catalogueUser, dataRow.name, dataRow.description, dataRow.sourceDataModelName,
                                         dataRow.targetDataModelName)

            dataFlow.validate()

            if (dataFlow.hasErrors()) {
                dataFlows += dataFlow
                return
            }

            addMetadataToCatalogueItem dataFlow, dataRow.metadata, catalogueUser

            if (!workbook.getSheet(dataRow.sheetKey)) {
                logger.warn('DataModel {} with key {} will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
            } else {

                List<DataFlowComponentDataRow> componentDataRows = loadDataRows(workbook, DataFlowComponentDataRow, filename,
                                                                                dataRow.sheetKey,
                                                                                DATAFLOWCOMPONENT_HEADER_ROWS,
                                                                                DATAFLOWCOMPONENT_ID_COLUMN)

                if (componentDataRows) {
                    dataFlows += loadSheetDataRowsIntoDataFlow(componentDataRows, dataFlow, catalogueUser)
                }
            }
        }

        dataFlows
    }

    DataFlow loadSheetDataRowsIntoDataFlow(List<DataFlowComponentDataRow> dataRows, DataFlow dataFlow, CatalogueUser catalogueUser) {

        dataRows.each {dataRow ->

            DataFlowComponent dataFlowComponent = dataFlowComponentService.createDataFlowComponent(catalogueUser, dataRow.transformLabel,
                                                                                                   dataRow.transformDescription)

            if (dataRow.mergedContentRows) {

                dataRow.mergedContentRows.each {mcr ->

                    if (mcr.hasSource()) {
                        DataElement sourceElement = dataElementService.findByDataClassPathAndLabel(dataFlow.source, mcr.sourceDataClassPathList,
                                                                                                   mcr.sourceDataElementName)
                        if (sourceElement) dataFlowComponent.addToSourceElements(sourceElement)
                    }

                    if (mcr.hasTarget()) {
                        DataElement targetElement = dataElementService.findByDataClassPathAndLabel(dataFlow.target, mcr.targetDataClassPathList,
                                                                                                   mcr.targetDataElementName)
                        if (targetElement) dataFlowComponent.addToTargetElements(targetElement)
                    }
                }

            } else {
                DataElement sourceElement = dataElementService.findByDataClassPathAndLabel(dataFlow.source, dataRow.sourceDataClassPathList,
                                                                                           dataRow.sourceDataElementName)
                DataElement targetElement = dataElementService.findByDataClassPathAndLabel(dataFlow.target, dataRow.targetDataClassPathList,
                                                                                           dataRow.targetDataElementName)
                if (targetElement) dataFlowComponent.addToTargetElements(targetElement)
                if (sourceElement) dataFlowComponent.addToSourceElements(sourceElement)
            }

            addMetadataToCatalogueItem dataFlowComponent, dataRow.metadata, catalogueUser

            dataFlow.addToDataFlowComponents dataFlowComponent
        }

        dataFlow
    }

    void addMetadataToCatalogueItem(CatalogueItem catalogueItem, List<MetadataColumn> metadata, CatalogueUser catalogueUser) {
        metadata.each {md ->
            logger.trace('Adding Metadata {}', md.key)
            catalogueItem.addToMetadata(md.namespace ?: namespace, md.key, md.value, catalogueUser)
        }
    }
}
