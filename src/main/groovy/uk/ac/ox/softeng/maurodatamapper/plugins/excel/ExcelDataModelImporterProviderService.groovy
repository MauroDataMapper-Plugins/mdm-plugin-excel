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
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.PrimitiveType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.PrimitiveTypeService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceType
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
import groovy.util.logging.Slf4j

@Slf4j
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

    private User currentUser

    @Override
    String getDisplayName() {
        'Excel (XLSX) Importer'
    }

    @Override
    String getVersion() {
        ExcelPlugin.VERSION
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
        this.currentUser = currentUser

        FileParameter importFile = importParameters.importFile
        if (!importFile.fileContents.size()) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')

        log.info('Importing {} as {}', importFile.fileName, currentUser.emailAddress)
        loadWorkbookFromInputStream(importFile).withCloseable { Workbook workbook ->
            loadDataModels(workbook, importFile.fileName).findAll { it.dataClasses }
        }
    }

    private static Collection<List<ContentDataRow>> loadDataClassRows(List<ContentDataRow> dataRows) {
        dataRows.groupBy { it.dataClassPath }.sort().values()
    }

    private static List<ContentDataRow> loadDataElementRows(List<ContentDataRow> dataRows) {
        dataRows.findAll { it.dataElementName }
    }

    private List<DataModel> loadDataModels(Workbook workbook, String filename) {
        List<DataModel> dataModels = []
        loadDataModelDataRows(workbook, filename).each { DataModelDataRow dataRow ->
            if (!workbook.getSheet(dataRow.sheetKey)) {
                getLog().warn('DataModel [{}] with key [{}] will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
                return []
            }

            List<ContentDataRow> contentDataRows = loadContentDataRows(workbook, filename, dataRow.sheetKey)
            if (!contentDataRows) return []
            dataModels << loadDataModel(dataRow, contentDataRows)
        }
        dataModels
    }

    private List<DataModelDataRow> loadDataModelDataRows(Workbook workbook, String filename) {
        log.info('Loading DataModel rows from {}', filename)
        loadDataRows(workbook, DataModelDataRow, filename, ExcelPlugin.DATAMODELS_SHEET_NAME, ExcelPlugin.DATAMODELS_NUM_HEADER_ROWS,
                     ExcelPlugin.DATAMODELS_ID_COLUMN_INDEX)
    }

    private List<ContentDataRow> loadContentDataRows(Workbook workbook, String filename, String sheetName) {
        log.info('Loading content (DataClass/DataElement) rows from {} in sheet {}', filename, sheetName)
        loadDataRows(workbook, ContentDataRow, filename, sheetName, ExcelPlugin.CONTENT_NUM_HEADER_ROWS, ExcelPlugin.CONTENT_ID_COLUMN_INDEX)
    }

    private DataModel loadDataModel(DataModelDataRow dataModelDataRow, List<ContentDataRow> contentDataRows) {
        log.info('Loading DataModel [{}]', dataModelDataRow.name)
        createDataModel(dataModelDataRow).tap { DataModel dataModel ->
            // Generate DataClasses in path-descending order, ensuring specific parent data is created before creating children
            getLog().info('Loading DataClass rows for DataModel [{}]', dataModel.label)
            loadDataClassRows(contentDataRows).each { findOrCreateDataClass(dataModel, it) }

            getLog().info('Loading DataElement rows for DataModel [{}]', dataModel.label)
            loadDataElementRows(contentDataRows).each { findOrCreateDataElement(dataModel, it) }
        }
    }

    private DataModel createDataModel(DataModelDataRow dataRow) {
        log.debug('Creating DataModel [{}]', dataRow.name)
        dataModelService.createAndSaveDataModel(
            currentUser, Folder.findOrCreateByLabel('random', [createdBy: currentUser.emailAddress]),
            DataModelType.findForLabel(dataRow.type), dataRow.name, dataRow.description, dataRow.author, dataRow.organisation,
            authorityService.defaultAuthority, saveDataModelsOnCreate).tap { DataModel dataModel ->
            addMetadataToCatalogueItem(dataModel, dataRow.metadata)
        }
    }

    private DataClass findOrCreateDataClass(DataModel dataModel, List<ContentDataRow> dataRows) {
        // Find the row with no DataElement name. If none exists, then just grab the first row.
        ContentDataRow dataRow = dataRows.find { !it.dataElementName } ?: dataRows.first()
        log.debug('Creating DataClass [{}]', dataRow.dataClassPath)
        dataClassService.findOrCreateDataClassByPath(
            dataModel, dataRow.dataClassPathList,
            dataRow.dataElementName ? null : dataRow.description,
            currentUser,
            dataRow.dataElementName ? 1 : dataRow.minMultiplicity,
            dataRow.dataElementName ? 1 : dataRow.maxMultiplicity).tap { DataClass dataClass ->
            if (!dataRow.dataElementName) addMetadataToCatalogueItem(dataClass, dataRow.metadata)
        }
    }

    private DataElement findOrCreateDataElement(DataModel dataModel, ContentDataRow dataRow) {
        DataClass dataClass = dataClassService.findDataClassByPath(dataModel, dataRow.dataClassPathList)
        DataType dataType = findOrCreateDataType(dataModel, dataRow)
        log.debug('Adding DataElement [{}] to DataClass [{}] with DataType [{}]', dataRow.dataElementName, dataClass.label, dataType.label)
        dataElementService.findOrCreateDataElementForDataClass(
            dataClass, dataRow.dataElementName, dataRow.description, currentUser, dataType, dataRow.minMultiplicity, dataRow.maxMultiplicity).tap {
            addMetadataToCatalogueItem(it, dataRow.metadata)
        }
    }

    private void addMetadataToCatalogueItem(CatalogueItem catalogueItem, List<MetadataColumn> metadata) {
        metadata.each { MetadataColumn metadataColumn ->
            getLog().debug('Adding Metadata [{}] to CatalogueItem [{}]', metadataColumn.key, catalogueItem.label)
            catalogueItem.addToMetadata(metadataColumn.namespace ?: namespace, metadataColumn.key, metadataColumn.value, currentUser)
        }
    }

    private DataType findOrCreateDataType(DataModel dataModel, ContentDataRow dataRow) {
        if (dataRow.mergedContentRows) return findOrCreateEnumerationType(dataModel, dataRow)
        if (dataRow.referenceToDataClassPath) return findOrCreateReferenceType(dataModel, dataRow)
        findOrCreatePrimitiveType(dataModel, dataRow)
    }

    private EnumerationType findOrCreateEnumerationType(DataModel dataModel, ContentDataRow dataRow) {
        log.debug('DataElement [{}] is an EnumerationType', dataRow.dataElementName)
        enumerationTypeService.findOrCreateDataTypeForDataModel(
            dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, currentUser).tap { EnumerationType enumerationType ->
            dataRow.mergedContentRows.each { enumerationType.addToEnumerationValues(it.key, it.value, currentUser as User) }
        }
    }

    private ReferenceType findOrCreateReferenceType(DataModel dataModel, ContentDataRow dataRow) {
        log.debug('DataElement [{}] is a ReferenceType', dataRow.dataElementName)
        DataClass referenceClass = dataClassService.findDataClassByPath(dataModel, dataRow.referenceToDataClassPathList)
        referenceTypeService.findOrCreateDataTypeForDataModel(
            dataModel, referenceClass.label, dataRow.dataTypeDescription, currentUser, referenceClass)
    }

    private PrimitiveType findOrCreatePrimitiveType(DataModel dataModel, ContentDataRow dataRow) {
        log.debug('DataElement [{}] is a PrimitiveType', dataRow.dataElementName)
        primitiveTypeService.findOrCreateDataTypeForDataModel(dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, currentUser)
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
