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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiBadRequestException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiUnauthorizedException
import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
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
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.exporter.ExcelDataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer.parameters.ExcelDataModelFileImporterProviderServiceParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.ContentDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.DataModelDataRow
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.MetadataColumn
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.workbook.WorkbookHandler
import uk.ac.ox.softeng.maurodatamapper.security.User

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Workbook

@Slf4j
@CompileStatic
class ExcelDataModelImporterProviderService extends DataModelImporterProviderService<ExcelDataModelFileImporterProviderServiceParameters>
    implements WorkbookHandler {

    DataClassService dataClassService
    DataElementService dataElementService
    PrimitiveTypeService primitiveTypeService
    EnumerationTypeService enumerationTypeService
    ReferenceTypeService referenceTypeService

    @Override
    String getDisplayName() {
        'Excel (XLSX) Importer'
    }

    @Override
    String getVersion() {
        getClass().getPackage().getSpecificationVersion() ?: 'SNAPSHOT'
    }

    @Override
    Boolean canImportMultipleDomains() {
        true
    }

    @Override
    String getNamespace() {
        'uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel'
    }

    @Override
    DataModel importModel(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importParameters) {
        importModels(currentUser, importParameters)?.first()
    }

    @Override
    List<DataModel> importModels(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EISP01', 'User must be logged in to import model')

        FileParameter importFile = importParameters.importFile
        if (!importFile.fileContents.size()) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')

        log.info('Importing {} as {}', importFile.fileName, currentUser.emailAddress)
        loadWorkbookFromInputStream(importFile).withCloseable {Workbook workbook ->
            loadDataModels(workbook, importFile.fileName, currentUser).findAll {it.dataClasses}
        }
    }

    private static Collection<List<ContentDataRow>> loadDataClassRows(List<ContentDataRow> dataRows) {
        dataRows.groupBy {it.dataClassPath}.sort().values()
    }

    private static List<ContentDataRow> loadDataElementRows(List<ContentDataRow> dataRows) {
        dataRows.findAll {it.dataElementName}
    }

    private List<DataModel> loadDataModels(Workbook workbook, String filename, User currentUser) {
        List<DataModel> dataModels = []
        loadDataModelDataRows(workbook, filename).each {DataModelDataRow dataRow ->
            if (!workbook.getSheet(dataRow.sheetKey)) {
                log.warn('DataModel [{}] with key [{}] will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
                return []
            }

            List<ContentDataRow> contentDataRows = loadContentDataRows(workbook, filename, dataRow.sheetKey)
            if (!contentDataRows) return []
            dataModels << loadDataModel(dataRow, contentDataRows, currentUser)
        }
        dataModels
    }

    private List<DataModelDataRow> loadDataModelDataRows(Workbook workbook, String filename) {
        log.info('Loading DataModel rows from {}', filename)
        loadDataRows(workbook, DataModelDataRow, filename, dataModelsSheetName,
                     ExcelDataModelExporterProviderService.DATAMODELS_NUM_HEADER_ROWS,
                     ExcelDataModelExporterProviderService.DATAMODELS_ID_COLUMN_INDEX)
    }

    private List<ContentDataRow> loadContentDataRows(Workbook workbook, String filename, String sheetName) {
        log.info('Loading content (DataClass/DataElement) rows from {} in sheet {}', filename, sheetName)
        loadDataRows(workbook, ContentDataRow, filename, sheetName,
                     ExcelDataModelExporterProviderService.CONTENT_NUM_HEADER_ROWS,
                     ExcelDataModelExporterProviderService.CONTENT_ID_COLUMN_INDEX)
    }

    private DataModel loadDataModel(DataModelDataRow dataModelDataRow, List<ContentDataRow> contentDataRows, User currentUser) {
        log.info('Loading DataModel [{}]', dataModelDataRow.name)
        createDataModel(dataModelDataRow, currentUser).tap {DataModel dataModel ->
            // Generate DataClasses in path-descending order, ensuring specific parent data is created before creating children
            getLog().info('Loading DataClass rows for DataModel [{}]', dataModel.label)
            loadDataClassRows(contentDataRows).each {findOrCreateDataClass(dataModel, it, currentUser)}

            getLog().info('Loading DataElement rows for DataModel [{}]', dataModel.label)
            loadDataElementRows(contentDataRows).each {findOrCreateDataElement(dataModel, it, currentUser)}
        }
    }

    private DataModel createDataModel(DataModelDataRow dataRow, User currentUser) {
        log.debug('Creating DataModel [{}]', dataRow.name)
        DataModel dataModel = new DataModel(label: dataRow.name,
                                            description: dataRow.description,
                                            organisation: dataRow.organisation,
                                            author: dataRow.author,
                                            type: DataModelType.findForLabel(dataRow.type),
                                            createdBy: currentUser.emailAddress
        )
        addMetadataToCatalogueItem(dataModel, dataRow.metadata, currentUser)
        dataModel
    }

    private DataClass findOrCreateDataClass(DataModel dataModel, List<ContentDataRow> dataRows, User currentUser) {
        // Find the row with no DataElement name. If none exists, then just grab the first row.
        ContentDataRow dataRow = dataRows.find {!it.dataElementName} ?: dataRows.first()
        log.debug('Creating DataClass [{}]', dataRow.dataClassPath)
        dataClassService.findOrCreateDataClassByPath(
            dataModel, dataRow.dataClassPathList,
            dataRow.dataElementName ? null : dataRow.description,
            currentUser,
            dataRow.dataElementName ? 1 : dataRow.minMultiplicity,
            dataRow.dataElementName ? 1 : dataRow.maxMultiplicity).tap {DataClass dataClass ->
            if (!dataRow.dataElementName) addMetadataToCatalogueItem(dataClass, dataRow.metadata, currentUser)
        }
    }

    private DataElement findOrCreateDataElement(DataModel dataModel, ContentDataRow dataRow, User currentUser) {
        DataClass dataClass = dataClassService.findDataClassByPath(dataModel, dataRow.dataClassPathList)
        DataType dataType = findOrCreateDataType(dataModel, dataRow, currentUser)
        log.debug('Adding DataElement [{}] to DataClass [{}] with DataType [{}]', dataRow.dataElementName, dataClass.label, dataType.label)
        dataElementService.findOrCreateDataElementForDataClass(
            dataClass, dataRow.dataElementName, dataRow.description, currentUser, dataType, dataRow.minMultiplicity, dataRow.maxMultiplicity).tap {
            addMetadataToCatalogueItem(it, dataRow.metadata, currentUser)
        }
    }

    private void addMetadataToCatalogueItem(CatalogueItem catalogueItem, List<MetadataColumn> metadata, User currentUser) {
        metadata.each {MetadataColumn metadataColumn ->
            getLog().debug('Adding Metadata [{}] to CatalogueItem [{}]', metadataColumn.key, catalogueItem.label)
            catalogueItem.addToMetadata(metadataColumn.namespace ?: namespace, metadataColumn.key, metadataColumn.value, currentUser)
        }
    }

    private DataType findOrCreateDataType(DataModel dataModel, ContentDataRow dataRow, User currentUser) {
        if (dataRow.mergedContentRows) return findOrCreateEnumerationType(dataModel, dataRow, currentUser)
        if (dataRow.referenceToDataClassPath) return findOrCreateReferenceType(dataModel, dataRow, currentUser)
        findOrCreatePrimitiveType(dataModel, dataRow, currentUser)
    }

    private EnumerationType findOrCreateEnumerationType(DataModel dataModel, ContentDataRow dataRow, User currentUser) {
        log.debug('DataElement [{}] is an EnumerationType', dataRow.dataElementName)
        enumerationTypeService.findOrCreateDataTypeForDataModel(
            dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, currentUser).tap {EnumerationType enumerationType ->
            dataRow.mergedContentRows.each {enumerationType.addToEnumerationValues(it.key, it.value, currentUser as User)}
        }
    }

    private ReferenceType findOrCreateReferenceType(DataModel dataModel, ContentDataRow dataRow, User currentUser) {
        log.debug('DataElement [{}] is a ReferenceType', dataRow.dataElementName)
        DataClass referenceClass = dataClassService.findDataClassByPath(dataModel, dataRow.referenceToDataClassPathList)
        referenceTypeService.findOrCreateDataTypeForDataModel(
            dataModel, referenceClass.label, dataRow.dataTypeDescription, currentUser, referenceClass)
    }

    private PrimitiveType findOrCreatePrimitiveType(DataModel dataModel, ContentDataRow dataRow, User currentUser) {
        log.debug('DataElement [{}] is a PrimitiveType', dataRow.dataElementName)
        primitiveTypeService.findOrCreateDataTypeForDataModel(dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription, currentUser)
    }
}
