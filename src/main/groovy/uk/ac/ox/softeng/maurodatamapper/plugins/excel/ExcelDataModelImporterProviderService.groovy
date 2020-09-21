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

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiBadRequestException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiUnauthorizedException
import uk.ac.ox.softeng.maurodatamapper.core.authority.AuthorityService
import uk.ac.ox.softeng.maurodatamapper.core.container.Folder
import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClassService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElementService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationTypeService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.PrimitiveTypeService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceTypeService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.catalogue.ContentDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.catalogue.DataModelDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column.MetadataColumn
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.util.WorkbookHandler
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.beans.factory.annotation.Autowired

import groovy.transform.CompileStatic

import static ExcelPlugin.CONTENT_HEADER_ROWS
import static ExcelPlugin.CONTENT_ID_COLUMN
import static ExcelPlugin.DATAMODELS_HEADER_ROWS
import static ExcelPlugin.DATAMODELS_ID_COLUMN
import static ExcelPlugin.DATAMODELS_SHEET_NAME

/**
 * @since 01/03/2018
 */
@CompileStatic
class ExcelDataModelImporterProviderService extends DataModelImporterProviderService<ExcelFileImporterProviderServiceParameters>
    implements WorkbookHandler {

    @Autowired
    DataModelService dataModelService

    @Autowired
    PrimitiveTypeService primitiveTypeService

    @Autowired
    DataClassService dataClassService

    @Autowired
    EnumerationTypeService enumerationTypeService

    @Autowired
    ReferenceTypeService referenceTypeService

    @Autowired
    DataElementService dataElementService

    @Autowired
    AuthorityService authorityService

    boolean saveDataModelsOnCreate = true

    @Override
    String getDisplayName() {
        'Excel (XLSX) Importer'
    }

    @Override
    String getVersion() {
        '1.0.0'
    }

    @Override
    Boolean canImportMultipleDomains() {
        true
    }

    @Override
    DataModel importDataModel(User currentUser, ExcelFileImporterProviderServiceParameters params) {
        importDataModels(currentUser, params)?.first()
    }

    @Override
    List<DataModel> importDataModels(User currentUser, ExcelFileImporterProviderServiceParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EISP01', 'User must be logged in to import model')
        if (importerParameters.importFile.fileContents.size() == 0) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')
        log.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)

        FileParameter importFile = importerParameters.importFile

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream)

            List<DataModelDataRow> dataModelDataRows = loadDataRows(workbook, DataModelDataRow, importFile.fileName,
                                                                    DATAMODELS_SHEET_NAME, DATAMODELS_HEADER_ROWS,
                                                                    DATAMODELS_ID_COLUMN)
            if (!dataModelDataRows) return []

            List<DataModel> dataModels = loadDataModels(dataModelDataRows, currentUser, workbook, importFile.fileName)

            return dataModels.findAll { it.dataClasses }
        } finally {
            closeWorkbook workbook
        }
    }

    List<DataModel> loadDataModels(List<DataModelDataRow> dataRows, User catalogueUser, Workbook workbook, String filename) {
        List<DataModel> dataModels = []

        dataRows.each { dataRow ->
            log.debug('Creating DataModel {}', dataRow.name)
            DataModel dataModel = dataModelService.createAndSaveDataModel(
                catalogueUser, Folder.findOrCreateByLabel('random', [createdBy: catalogueUser.emailAddress]),
                DataModelType.findForLabel(dataRow.type), dataRow.name, dataRow.description, dataRow.author, dataRow.organisation,
                authorityService.defaultAuthority, saveDataModelsOnCreate)

            addMetadataToCatalogueItem dataModel, dataRow.metadata, catalogueUser

            if (!workbook.getSheet(dataRow.sheetKey)) {
                log.warn('DataModel {} with key {} will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
                return
            }

            List<ContentDataRow> contentDataRows = loadDataRows(workbook, ContentDataRow, filename, dataRow.sheetKey,
                                                                CONTENT_HEADER_ROWS, CONTENT_ID_COLUMN)

            if (!contentDataRows) return

            dataModels += loadSheetDataRowsIntoDataModel(contentDataRows, dataModel, catalogueUser)
        }

        dataModels
    }

    DataModel loadSheetDataRowsIntoDataModel(List<ContentDataRow> dataRows, DataModel dataModel, User catalogueUser) {

        // Load DataClasses
        loadDataClassRows(dataRows, dataModel, catalogueUser)

        // Load the rest of the rows as DataElements
        loadDataElementRows(dataRows.findAll { it.dataElementName }, dataModel, catalogueUser)

        dataModel
    }

    void loadDataClassRows(List<ContentDataRow> dataRows, DataModel dataModel, User catalogueUser) {
        log.info('Loading DataClass rows for DataModel {}', dataModel.label)

        // Where not DataElement name then the row is DataClass specific data, otherwise we just need to make sure all the DataClasses exist
        // Sort the rows so we generate them in path descending order, this will ensure we create any specific data before we create children
        dataRows.groupBy { it.dataClassPath }.sort().each { path, dataClassRows ->

            // Find the row with no DataElement name, if it doesn't exist then just grab the first row
            ContentDataRow dataRow = dataClassRows.find { !it.dataElementName } ?: dataClassRows.first()

            String description = dataRow.dataElementName ? null : dataRow.description
            Integer min = dataRow.dataElementName ? 1 : dataRow.minMultiplicity
            Integer max = dataRow.dataElementName ? 1 : dataRow.maxMultiplicity

            log.debug('Creating DataClass [{}]', dataRow.dataClassPath)
            // TODO(adjl): Compose this functinoality from other methods instead of publicising findOrCreateDataClassByPath()
            DataClass dataClass = dataClassService.findOrCreateDataClassByPath(
                dataModel, dataRow.dataClassPathList, description, catalogueUser, min, max)
            if (!dataRow.dataElementName) {
                addMetadataToCatalogueItem dataClass, dataRow.metadata, catalogueUser
            }
        }
    }

    void loadDataElementRows(List<ContentDataRow> dataRows, DataModel dataModel, User catalogueUser) {
        dataRows.each { dataRow ->

            DataClass parentDataClass = dataClassService.findDataClassByPath(dataModel, dataRow.dataClassPathList)
            DataType dataType = null

            if (dataRow.mergedContentRows) {
                log.debug('DataElement is an EnumerationType')

                dataType = enumerationTypeService.findOrCreateDataTypeForDataModel(dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription,
                                                                                   catalogueUser)

                dataRow.mergedContentRows.each { mcr ->
                    ((EnumerationType) dataType).addToEnumerationValues(mcr.key, mcr.value, catalogueUser)
                }
            } else if (dataRow.referenceToDataClassPath) {
                DataClass referenceClass = dataClassService.findDataClassByPath(dataModel, dataRow.referenceToDataClassPathList)
                dataType = referenceTypeService.findOrCreateDataTypeForDataModel(dataModel, referenceClass.label, dataRow.dataTypeDescription,
                                                                                 catalogueUser, referenceClass)
            } else {
                dataType = primitiveTypeService.findOrCreateDataTypeForDataModel(
                    dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, catalogueUser)
            }

            log.debug('Adding DataElement [{}] to DataClass [{}] with DataType [{}]', dataRow.dataElementName, parentDataClass?.label,
                      dataType?.label)
            DataElement dataElement = dataElementService.findOrCreateDataElementForDataClass(
                parentDataClass, dataRow.dataElementName, dataRow.description, catalogueUser, dataType, dataRow.minMultiplicity,
                dataRow.maxMultiplicity
            )

            addMetadataToCatalogueItem dataElement, dataRow.metadata, catalogueUser
        }

        dataModel
    }

    void addMetadataToCatalogueItem(CatalogueItem catalogueItem, List<MetadataColumn> metadata, User catalogueUser) {
        metadata.each { md ->
            log.trace('Adding Metadata {}', md.key)
            catalogueItem.addToMetadata(md.namespace ?: namespace, md.key, md.value, catalogueUser)
        }
    }

    @Override
    // TODO(adjl): Investigate why casting `as User` in DataModelImporterProviderService doesn't work
    DataModel importDomain(User currentUser, ExcelFileImporterProviderServiceParameters params) {
        DataModel dataModel = importDataModel(currentUser, params)
        if (!dataModel) return null
        if (params.modelName) dataModel.label = params.modelName
        checkImport(currentUser as User, dataModel, params.finalised, params.importAsNewDocumentationVersion)
    }

    @Override
    // TODO(adjl): Investigate why checkImport doesn't work in DataModelImporterProviderService despite matching param types
    List<DataModel> importDomains(User currentUser, ExcelFileImporterProviderServiceParameters params) {
        List<DataModel> dataModels = importDataModels(currentUser, params)
        dataModels?.collect { checkImport(currentUser, it, params.finalised, params.importAsNewDocumentationVersion) }
    }
}
