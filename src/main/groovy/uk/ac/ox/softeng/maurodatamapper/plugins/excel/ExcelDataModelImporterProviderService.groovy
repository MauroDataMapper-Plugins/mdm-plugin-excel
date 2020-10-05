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
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.ContentDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.DataModelDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.MetadataColumn
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.workbook.WorkbookHandler
import uk.ac.ox.softeng.maurodatamapper.security.User

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.beans.factory.annotation.Autowired

import groovy.transform.CompileStatic

// @Slf4j
@CompileStatic
class ExcelDataModelImporterProviderService extends DataModelImporterProviderService<ExcelDataModelFileImporterProviderServiceParameters>
    implements WorkbookHandler {

    @Autowired
    DataModelService dataModelService

    @Autowired
    AuthorityService authorityService

    @Autowired
    DataClassService dataClassService

    @Autowired
    DataElementService dataElementService

    @Autowired
    PrimitiveTypeService primitiveTypeService

    @Autowired
    EnumerationTypeService enumerationTypeService

    @Autowired
    ReferenceTypeService referenceTypeService

    boolean saveDataModelsOnCreate = true

    @Override
    String getDisplayName() {
        'Excel (XLSX) Importer'
    }

    @Override
    String getVersion() {
        '2.0.0-SNAPSHOT'
    }

    @Override
    Boolean canImportMultipleDomains() {
        true
    }

    @Override
    DataModel importDataModel(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importParameters) {
        importDataModels(currentUser, importParameters)?.first()
    }

    @Override
    List<DataModel> importDataModels(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EISP01', 'User must be logged in to import model')

        FileParameter importFile = importParameters.importFile
        if (!importFile.fileContents.size()) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')
        getLog().info('Importing {} as {}', importFile.fileName, currentUser.emailAddress)

        Workbook workbook = null
        try {
            workbook = loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream)
            List<DataModelDataRow> dataRows = loadDataRows(
                workbook, DataModelDataRow, importFile.fileName, ExcelPlugin.DATAMODELS_SHEET_NAME, ExcelPlugin.DATAMODELS_HEADER_ROWS,
                ExcelPlugin.DATAMODELS_ID_COLUMN)
            if (!dataRows) return []
            loadDataModels(dataRows, currentUser, workbook, importFile.fileName).findAll { it.dataClasses }
        } finally {
            closeWorkbook(workbook)
        }
    }

    private List<DataModel> loadDataModels(List<DataModelDataRow> dataRows, User currentUser, Workbook workbook, String filename) {
        List<DataModel> dataModels = []
        dataRows.each { DataModelDataRow dataRow ->
            getLog().debug('Creating DataModel {}', dataRow.name)
            DataModel dataModel = dataModelService.createAndSaveDataModel(
                currentUser, Folder.findOrCreateByLabel('random', [createdBy: currentUser.emailAddress]),
                DataModelType.findForLabel(dataRow.type), dataRow.name, dataRow.description, dataRow.author, dataRow.organisation,
                authorityService.defaultAuthority, saveDataModelsOnCreate)
            addMetadataToCatalogueItem(dataModel, dataRow.metadata, currentUser)

            if (!workbook.getSheet(dataRow.sheetKey)) {
                getLog().warn('DataModel {} with key {} will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
                return
            }

            List<ContentDataRow> contentDataRows = loadDataRows(
                workbook, ContentDataRow, filename, dataRow.sheetKey, ExcelPlugin.CONTENT_HEADER_ROWS, ExcelPlugin.CONTENT_ID_COLUMN)
            if (!contentDataRows) return
            dataModels << loadSheetDataRowsIntoDataModel(contentDataRows, dataModel, currentUser)
        }
        dataModels
    }

    private DataModel loadSheetDataRowsIntoDataModel(List<ContentDataRow> dataRows, DataModel dataModel, User currentUser) {
        loadDataClassRows(dataRows, dataModel, currentUser) // Load DataClasses
        loadDataElementRows(dataRows.findAll { it.dataElementName }, dataModel, currentUser) // Load the rest of the rows as DataElements
        dataModel
    }

    private void loadDataClassRows(List<ContentDataRow> dataRows, DataModel dataModel, User currentUser) {
        getLog().info('Loading DataClass rows for DataModel {}', dataModel.label)
        // If the DataElement name is empty, then the row is DataClass-specific data. Otherwise, we only need to ensure all the DataClasses exist.
        // Sort the rows so we generate them in path-descending order. This ensures we create any specific data before we create children.
        dataRows.groupBy { it.dataClassPath }.sort().each { String path, List<ContentDataRow> dataClassRows ->
            // Find the row with no DataElement name. If none exists, then just grab the first row.
            ContentDataRow dataRow = dataClassRows.find { !it.dataElementName } ?: dataClassRows.first()
            getLog().debug('Creating DataClass [{}]', dataRow.dataClassPath)
            // TODO: Compose this functionality from other methods instead of publicising findOrCreateDataClassByPath()
            DataClass dataClass = dataClassService.findOrCreateDataClassByPath(
                dataModel, dataRow.dataClassPathList,
                dataRow.dataElementName ? null : dataRow.description,
                currentUser,
                dataRow.dataElementName ? 1 : dataRow.minMultiplicity,
                dataRow.dataElementName ? 1 : dataRow.maxMultiplicity)
            if (!dataRow.dataElementName) {
                addMetadataToCatalogueItem(dataClass, dataRow.metadata, currentUser)
            }
        }
    }

    private void loadDataElementRows(List<ContentDataRow> dataRows, DataModel dataModel, User currentUser) {
        dataRows.each { ContentDataRow dataRow ->
            DataType dataType
            if (dataRow.mergedContentRows) {
                getLog().debug('DataElement is an EnumerationType')
                dataType = enumerationTypeService.findOrCreateDataTypeForDataModel(
                    dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, currentUser)
                dataRow.mergedContentRows.each { (dataType as EnumerationType).addToEnumerationValues(it.key, it.value, currentUser) }
            } else if (dataRow.referenceToDataClassPath) {
                DataClass referenceClass = dataClassService.findDataClassByPath(dataModel, dataRow.referenceToDataClassPathList)
                dataType = referenceTypeService.findOrCreateDataTypeForDataModel(
                    dataModel, referenceClass.label, dataRow.dataTypeDescription, currentUser, referenceClass)
            } else {
                dataType = primitiveTypeService.findOrCreateDataTypeForDataModel(
                    dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, currentUser)
            }

            DataClass parentDataClass = dataClassService.findDataClassByPath(dataModel, dataRow.dataClassPathList)
            getLog().debug('Adding DataElement [{}] to DataClass [{}] with DataType [{}]', dataRow.dataElementName, parentDataClass?.label,
                           dataType?.label)
            DataElement dataElement = dataElementService.findOrCreateDataElementForDataClass(
                parentDataClass, dataRow.dataElementName, dataRow.description, currentUser, dataType, dataRow.minMultiplicity,
                dataRow.maxMultiplicity)
            addMetadataToCatalogueItem(dataElement, dataRow.metadata, currentUser)
        }
        dataModel
    }

    private void addMetadataToCatalogueItem(CatalogueItem catalogueItem, List<MetadataColumn> metadata, User currentUser) {
        metadata.each { MetadataColumn metadataColumn ->
            getLog().trace('Adding Metadata {}', metadataColumn.key)
            catalogueItem.addToMetadata(metadataColumn.namespace ?: namespace, metadataColumn.key, metadataColumn.value, currentUser)
        }
    }

    // The following are hacks– exact reimplementations of methods inherited from DataModelImporterProviderService. Despite having little to no
    // substantial differences in the code, simply calling importDomain() and importDomains()––prior to reimplementing––results in a build error
    // for some reason. But re-overriding them does, hence the code duplication below. Requires deeper investigation.
    //
    // TODO: Investigate why...
    // - Casting currentUser `as User` in importDomain() does not work
    // - Calling checkImport() does not work despite matching parameter types

    @Override
    DataModel importDomain(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importParameters) {
        DataModel dataModel = importDataModel(currentUser, importParameters)
        if (!dataModel) return null
        if (importParameters.modelName) dataModel.label = importParameters.modelName
        checkImport(currentUser as User, dataModel, importParameters.finalised, importParameters.importAsNewDocumentationVersion)
    }

    @Override
    List<DataModel> importDomains(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importParameters) {
        List<DataModel> dataModels = importDataModels(currentUser, importParameters)
        dataModels?.collect { checkImport(currentUser, it, importParameters.finalised, importParameters.importAsNewDocumentationVersion) }
    }
}
