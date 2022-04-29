/*
 * Copyright 2020-2022 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiInternalException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiUnauthorizedException
import uk.ac.ox.softeng.maurodatamapper.core.facet.Metadata
import uk.ac.ox.softeng.maurodatamapper.core.model.facet.MetadataAware
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClassService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElementService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.PrimitiveType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceTypeService
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.enumeration.EnumerationValue
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer.parameters.ExcelDataModelFileImporterProviderServiceParameters
import uk.ac.ox.softeng.maurodatamapper.security.User

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.apache.poi.EncryptedDocumentException
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

/**
 * @since 01/03/2018
 */
@Slf4j
@CompileStatic
class SimpleExcelDataModelImporterProviderService
    extends DataModelImporterProviderService<ExcelDataModelFileImporterProviderServiceParameters> {

    static final Map<String, String> DATAMODEL_SHEET_COLUMNS = [name: 'Name',
                                                                description: 'Description',
                                                                author: 'Author',
                                                                organisation: 'Organisation',
                                                                sheetKey: 'Sheet Key',
                                                                type: 'Type']
    static final Map<String, String> ENUM_SHEET_COLUMNS = [dataModelName: 'DataModel Name',
                                                           enumerationName: 'Enumeration Name',
                                                           description: 'Description',
                                                           key: 'Key',
                                                           value: 'Value']
    static final Map<String, String> MODEL_SHEET_COLUMNS = [dataClassPath  : 'DataClass Path',
                                                            dataElementName: 'DataElement Name',
                                                            description: 'Description',
                                                            minMultiplicity: 'Minimum[ \r\n]+Multiplicity',
                                                            maxMultiplicity: 'Maximum[ \r\n]+Multiplicity',
                                                            dataTypeName   : 'DataType Name',
                                                            dataTypeReference: 'DataType Reference']
    static DataFormatter dataFormatter = new DataFormatter()

    DataClassService dataClassService
    ReferenceTypeService referenceTypeService
    DataElementService dataElementService

    @Override
    String getDisplayName() {
        'Simple Excel (XLSX) Importer'
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
    DataModel importModel(User currentUser, ExcelDataModelFileImporterProviderServiceParameters params) {
        importModels(currentUser, params)?.first()
    }

    @Override
    List<DataModel> importModels(User currentUser, ExcelDataModelFileImporterProviderServiceParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EISP01', 'User must be logged in to import model')
        if (importerParameters.importFile.fileContents.size() == 0) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')
        log.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)

        FileParameter importFile = importerParameters.importFile

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream)

            Sheet dataModelsSheet = workbook.getSheet('DataModels')
            Sheet enumerationsSheet = workbook.getSheet('Enumerations')

            if (!dataModelsSheet) {
                throw new ApiInternalException('EFS02', 'The Excel file does not include a sheet called [DataModels]')
            }
            Map<String, Map<String, EnumerationType>> enumerationTypes = [:]
            if (enumerationsSheet) {
                enumerationTypes = calculateEnumerationTypes(currentUser, getSheetValues(ENUM_SHEET_COLUMNS, enumerationsSheet))
            }

            List<DataModel> dataModels = []
            List<Map<String, String>> sheetValues = getSheetValues(DATAMODEL_SHEET_COLUMNS, dataModelsSheet)
            sheetValues.each {row ->
                DataModel dataModel = dataModelFromRow(currentUser, row)
                addMetadataFromExtraColumns(dataModel, DATAMODEL_SHEET_COLUMNS, row)
                String sheetKey = row.sheetKey
                Sheet modelSheet = workbook.getSheet(sheetKey)
                if (!dataModelsSheet) {
                    throw new ApiInternalException('EFS06', "The Excel file does not include a sheet called ${sheetKey} referenced for " +
                                                            "model'${dataModel.label}'")
                }
                addClassesAndElements(currentUser, dataModel, modelSheet, enumerationTypes[dataModel.label] ?: [:] as Map<String,EnumerationType> )

                dataModels.add(dataModel)
            }

            return dataModels
        } finally {
            closeWorkbook workbook
        }
    }

    Workbook loadWorkbookFromFilename(String filename) throws ApiException {
        loadWorkbookFromInputStream filename, getClass().classLoader.getResourceAsStream(filename)
    }

    Workbook loadWorkbookFromInputStream(String filename, InputStream inputStream) throws ApiException {
        if (!inputStream) throw new ApiInternalException('EFS01', "No inputstream for ${filename}")
        try {
            ZipSecureFile.setMinInflateRatio(0);
            return WorkbookFactory.create(inputStream)
        } catch (EncryptedDocumentException ignored) {
            throw new ApiInternalException('EFS02', "Excel file ${filename} could not be read as it is encrypted")
        } catch (InvalidFormatException ignored) {
            throw new ApiInternalException('EFS03', "Excel file ${filename} could not be read as it is not a valid format")
        } catch (IOException ex) {
            throw new ApiInternalException('EFS04', "Excel file ${filename} could not be read", ex)
        }
    }

    void closeWorkbook(Workbook workbook) {
        if (workbook != null) {
            try {
                workbook.close()
            } catch (IOException ignored) {
                // ignored
            }
        }
    }

    List<Map<String, String>> getSheetValues(Map<String,String> intendedColumnMappings, Sheet sheet) {
        List<Map<String, String>> returnValues = []
        Map<String, Integer> expectedSheetColumns = [:]
        Map<String, Integer> otherSheetColumns = [:]
        Row row = sheet.getRow(0)
        int col = 0
        while (row.getCell(col)) {
            String headerText = getCellValueAsString(row.getCell(col))
            boolean found = false
            intendedColumnMappings.each {columnName, regex ->
                if (headerText.toLowerCase().trim() ==~regex.toLowerCase().trim()) {
                    expectedSheetColumns[columnName] = col
                    found = true
                }
            }
            if (!found) {
                otherSheetColumns[headerText] = col
            }
            col++
        }
        if (expectedSheetColumns.size() != intendedColumnMappings.size()) {
            throw new ApiBadRequestException('EIS03',
                                             "Missing header: ${sheet.getSheetName()} sheet should include the following headers: ${intendedColumnMappings.values()}")
        }

        Iterator<Row> rowIterator = sheet.rowIterator()
        // burn the header row
        rowIterator.next()

        while (rowIterator.hasNext()) {
            row = rowIterator.next()
            Map<String, String> rowValues = [:]
            intendedColumnMappings.each {columnName,regex ->
                String value = getCellValueAsString(row.getCell(expectedSheetColumns[columnName]))
                rowValues[columnName] = value
            }
            otherSheetColumns.keySet().each {columnName ->
                String value = getCellValueAsString(row.getCell(otherSheetColumns[columnName]))
                rowValues[columnName] = value
            }
            returnValues.add(rowValues)
        }
        return returnValues
    }

    void addMetadataFromExtraColumns(MetadataAware entity, Map<String,String> expectedColumnMappings, Map<String, String> columnValues) {
        Set<String> expectedKeys = expectedColumnMappings.keySet()
        columnValues.keySet().each {columnName ->
            if (!expectedKeys.contains(columnName)) {
                String key = columnName
                String namespace = this.getNamespace()
                if (key.contains(':')) {
                    String[] components = key.split(':')
                    namespace = components[0]
                    key = components[1]
                }
                if (columnValues[columnName]) {
                    entity.addToMetadata(new Metadata(namespace: namespace, key: key, value: columnValues[columnName]))
                }
            }
        }
    }

    DataModel dataModelFromRow(User currentUser, Map<String, String> columnValues) {
        String label = columnValues.name
        String description = columnValues.description
        String author = columnValues.author
        String organisation = columnValues.organisation
        String type = columnValues.type
        DataModelType dataModelType
        if (type.toLowerCase().replaceAll('[ _]', '') == 'dataasset') {
            dataModelType = DataModelType.DATA_ASSET
        } else if (type.toLowerCase().replaceAll('[ _]', '') == 'datastandard') {
            dataModelType = DataModelType.DATA_STANDARD
        } else {
            throw new ApiBadRequestException('SEIS03', "Invalid Data Model Type for Model '${label}'")
        }

        new DataModel(label: label,
                      description: description,
                      organisation: organisation,
                      author: author,
                      type: DataModelType.findForLabel(dataModelType.toString()),
                      createdBy: currentUser.emailAddress)
    }

    Map<String, Map<String, EnumerationType>> calculateEnumerationTypes(User createdBy, List<Map<String, String>> sheetValues) {

        Map<String, Map<String, EnumerationType>> returnValues = [:]
        sheetValues.each {columnValues ->
            String dataModelName = columnValues.dataModelName
            String label = columnValues.enumerationName
            String description = columnValues.description
            String key = columnValues.key
            String value = columnValues.value

            Map<String, EnumerationType> modelEnumTypes = returnValues[dataModelName]
            if (!modelEnumTypes) {
                modelEnumTypes = [:]
                returnValues[dataModelName] = modelEnumTypes
            }
            EnumerationType enumerationType = modelEnumTypes[label]
            if (!enumerationType) {
                enumerationType = new EnumerationType(createdBy: createdBy.emailAddress, label: label, description: description)
                modelEnumTypes[label] = enumerationType
            }
            EnumerationValue enumerationValue = new EnumerationValue(createdBy: createdBy.emailAddress, key: key, value: value)
            addMetadataFromExtraColumns(enumerationValue, ENUM_SHEET_COLUMNS, columnValues)
            enumerationType.addToEnumerationValues(enumerationValue)
        }
        return returnValues
    }

    void addClassesAndElements(User currentUser, DataModel dataModel, Sheet dataModelSheet, Map<String, EnumerationType> enumerationTypes) {
        List<Map<String, String>> sheetValues = getSheetValues(MODEL_SHEET_COLUMNS, dataModelSheet)
        List<Map<String, String>> referenceTypeDataElementRows = sheetValues.findAll {it.dataTypeReference}
        Map<String, DataType> modelDataTypes = [:]
        DataClass dataClass
        String previousDataClassPath
        (sheetValues - referenceTypeDataElementRows).each {row ->
            String dataClassPath = row.dataClassPath
            String name = row.dataElementName
            MetadataAware createdElement = null

            if (!previousDataClassPath?.equals(dataClassPath) && name || !name) {
                // We're dealing with a data class
                previousDataClassPath = dataClassPath
                String description = row.description
                String minMult = row.minMultiplicity
                String maxMult = row.maxMultiplicity
                dataClass = getOrCreateClassFromPath(currentUser, dataModel, dataClassPath, name ? '' : description,
                                                     name ? 1 : minMult ? Integer.parseInt(minMult) : 1,
                                                     name ? 1 : maxMult ? maxMult == '*' ? -1 : Integer.parseInt(maxMult) : 1)
                createdElement = name ? addDataElement(currentUser, dataModel, dataClass, modelDataTypes, enumerationTypes, row) : dataClass
            } else {
                createdElement = addDataElement(currentUser, dataModel, dataClass, modelDataTypes, enumerationTypes, row)
            }
            addMetadataFromExtraColumns(createdElement, MODEL_SHEET_COLUMNS, row)
        }
        referenceTypeDataElementRows.each {row ->
            String dataClassPath = row.dataClassPath
            String description = row.description
            String minMult = row.minMultiplicity
            String maxMult = row.maxMultiplicity
            MetadataAware createdElement = null
            dataClass = getOrCreateClassFromPath(currentUser, dataModel, dataClassPath, description,
                                                 minMult ? Integer.parseInt(minMult) : 1,
                                                 maxMult ? maxMult == '*' ? -1 : Integer.parseInt(maxMult) : 1)
            createdElement = addDataElement(currentUser, dataModel, dataClass, modelDataTypes, enumerationTypes, row)
            addMetadataFromExtraColumns(createdElement, MODEL_SHEET_COLUMNS, row)
        }
        modelDataTypes.values().each {
            dataModel.addToDataTypes(it)
        }
        enumerationTypes.values().each {
            dataModel.addToDataTypes(it)
        }
    }

    DataClass getOrCreateClassFromPath(User currentUser, DataModel dataModel, String path, String description = '', Integer minMultiplicity = null,
                                       Integer maxMultiplicity = null) {
        dataClassService.findOrCreateDataClassByPath(
            dataModel, getDataClassPathLabels(path), description, currentUser, minMultiplicity, maxMultiplicity)
    }

    String getCellValueAsString(Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    private DataElement addDataElement(User currentUser, DataModel dataModel, DataClass parentDataClass, Map<String, DataType> modelDataTypes,
                                       Map<String, EnumerationType> enumerationTypes, Map<String, String> row) {
        String name = row.dataElementName
        String description = row.description
        String minMult = row.minMultiplicity
        String maxMult = row.maxMultiplicity
        String typeName = row.dataTypeName
        String typeReference = row.dataTypeReference
        // We're dealing with a data element
        DataType elementDataType
        if (enumerationTypes[typeName]) {
            elementDataType = enumerationTypes[typeName]
        } else if (modelDataTypes[typeName]) {
            elementDataType = modelDataTypes[typeName]
        } else {
            if (typeReference) {
                DataClass referenceDataClass = dataClassService.findDataClassByPath(dataModel, getDataClassPathLabels(typeReference))
                elementDataType = referenceTypeService.findOrCreateDataTypeForDataModel(
                    dataModel, referenceDataClass.label, null, currentUser, referenceDataClass)
            } else {
                elementDataType = new PrimitiveType(createdBy: currentUser.emailAddress, label: typeName)
            }
            modelDataTypes[typeName] = elementDataType
        }
        dataElementService.findOrCreateDataElementForDataClass(
            parentDataClass, name, description, currentUser, elementDataType, minMult ? Integer.parseInt(minMult) : 1,
            maxMult ? maxMult == '*' ? -1 : Integer.parseInt(maxMult) : 1)
    }

    private List<String> getDataClassPathLabels(String dataClassPath) {
        dataClassPath.split('\\|').toList()*.trim()
    }
}
