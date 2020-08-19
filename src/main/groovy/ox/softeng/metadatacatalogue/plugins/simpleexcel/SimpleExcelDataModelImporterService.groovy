package ox.softeng.metadatacatalogue.plugins.simpleexcel

import org.apache.poi.openxml4j.util.ZipSecureFile
import ox.softeng.metadatacatalogue.core.api.exception.ApiBadRequestException
import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.api.exception.ApiInternalException
import ox.softeng.metadatacatalogue.core.api.exception.ApiUnauthorizedException
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClassService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElement
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElementService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.DataType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationValue
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.PrimitiveType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.PrimitiveTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.ReferenceType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.ReferenceTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.facet.Metadata
import ox.softeng.metadatacatalogue.core.spi.importer.DataModelImporterPlugin
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.FileParameter
import ox.softeng.metadatacatalogue.core.traits.domain.MetadataAware
import ox.softeng.metadatacatalogue.core.type.catalogue.DataModelType
import ox.softeng.metadatacatalogue.core.user.CatalogueUser
import ox.softeng.metadatacatalogue.plugins.simpleexcel.parameter.SimpleExcelFileImporterParameters

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.springframework.beans.factory.annotation.Autowired

/**
 * @since 01/03/2018
 */
class SimpleExcelDataModelImporterService implements DataModelImporterPlugin<SimpleExcelFileImporterParameters> {

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

    static DataFormatter dataFormatter = new DataFormatter()

    @Override
    String getDisplayName() {
        'Simple Excel (XLSX) Importer'
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
    DataModel importDataModel(CatalogueUser currentUser, SimpleExcelFileImporterParameters params) {
        importDataModels(currentUser, params)?.first()
    }

    @Override
    List<DataModel> importDataModels(CatalogueUser currentUser, SimpleExcelFileImporterParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EISP01', 'User must be logged in to import model')
        if (importerParameters.importFile.fileContents.size() == 0) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')
        logger.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)

        FileParameter importFile = importerParameters.importFile

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream)

            Sheet dataModelsSheet = workbook.getSheet("DataModels")
            Sheet enumerationsSheet = workbook.getSheet("Enumerations")

            if(!dataModelsSheet) {
                throw new ApiInternalException('EFS02', "The Excel file does not include a sheet called 'DataModels'")
            }
            Map<String, DataType> dataTypes = [:]
            List<Map<String, String>> sheetValues =[]
            Map<String, Map<String, EnumerationType>> enumerationTypes = [:]
            if(enumerationsSheet) {
                enumerationTypes = calculateEnumerationTypes(sheetValues)
            }

            List<DataModel> dataModels = []
            sheetValues = getSheetValues(PluginSimpleExcelModule.DATAMODEL_SHEET_COLUMNS, dataModelsSheet)
            sheetValues.each {row ->
                DataModel dataModel = dataModelFromRow(row)
                addMetadataFromExtraColumns(dataModel, PluginSimpleExcelModule.DATAMODEL_SHEET_COLUMNS, row)
                String sheetKey = row["Sheet Key"]
                Sheet modelSheet = workbook.getSheet(sheetKey)
                if(!dataModelsSheet) {
                    throw new ApiInternalException('EFS06', "The Excel file does not include a sheet called ${sheetKey} referenced for " +
                                                            "model'${dataModel.label}'")
                }
                addClassesAndElements(dataModel, modelSheet, enumerationTypes[dataModel.label]?:[:])

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

    List<Map<String, String>> getSheetValues(List<String> intendedColumns, Sheet sheet) {
        List<Map<String, String>> returnValues = []
        Map<String, Integer> expectedSheetColumns = [:]
        Map<String, Integer> otherSheetColumns = [:]
        Row row = sheet.getRow(0)
        int col = 0
        while (row.getCell(col)) {
            String headerText = getCellValueAsString(row.getCell(col))
            boolean found = false
            intendedColumns.each {columnName ->
                if (headerText.toLowerCase().trim() == columnName.toLowerCase().trim()) {
                    expectedSheetColumns[columnName] = col
                    found = true
                }
            }
            if (!found) {
                otherSheetColumns[headerText] = col
            }
            col++
        }
        if (expectedSheetColumns.size() != intendedColumns.size()) {
            throw new ApiBadRequestException('EIS03',
                                             "Missing header: ${sheet.getSheetName()} sheet should include the following headers: ${intendedColumns}")
        }

        Iterator<Row> rowIterator = sheet.rowIterator()
        // burn the header row
        rowIterator.next()

        while (rowIterator.hasNext()) {
            row = rowIterator.next()
            Map<String, String> rowValues = [:]
            intendedColumns.each {columnName ->
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

    void addMetadataFromExtraColumns(MetadataAware entity, List<String> expectedColumns, Map<String, String> columnValues)
    {
        columnValues.keySet().each { columnName ->
            if(!expectedColumns.contains(columnName)) {
                String key = columnName
                String namespace = this.getNamespace()
                if(key.contains(":")) {
                    String[] components = key.split(":")
                    namespace = components[0]
                    key = components[1]
                }
                entity.addToMetadata(new Metadata(namespace: namespace, key: key,
                                                     value: columnValues[columnName]))
            }
        }
    }

    DataModel dataModelFromRow(Map<String, String> columnValues) {
        String label = columnValues["Name"]
        String description = columnValues["Description"]
        String author = columnValues["Author"]
        String organisation = columnValues["Organisation"]
        //String sheetKey = columnValues["Sheet Key"]
        String type = columnValues["Type"]
        DataModelType dataModelType
        if(type.toLowerCase().replaceAll("[ _]","") == "dataasset") {
            dataModelType = DataModelType.DATA_ASSET
        } else if(type.toLowerCase().replaceAll("[ _]","") == "datastandard") {
            dataModelType = DataModelType.DATA_STANDARD
        } else {
            throw new ApiBadRequestException('SEIS03', "Invalid Data Model Type for Model '${label}'")
        }

        DataModel dataModel = new DataModel(label: label,
                                            description: description,
                                            author: author,
                                            organisation: organisation,
                                            type: dataModelType)

        return dataModel
    }



    Map<String, Map<String, EnumerationType>> calculateEnumerationTypes(List<Map<String, String>> sheetValues) {

        Map<String, Map<String, EnumerationType>> returnValues = [:]
        sheetValues.each {columnValues ->
            String dataModelName = columnValues["DataModel Name"]
            String label = columnValues["Name"]
            String description = columnValues["Description"]
            String key = columnValues["Key"]
            String value = columnValues["Value"]

            Map<String, EnumerationType> modelEnumTypes = returnValues[dataModelName]
            if (!modelEnumTypes) {
                modelEnumTypes = [:]
                returnValues[dataModelName] = modelEnumTypes
            }
            EnumerationType enumerationType = modelEnumTypes[label]
            if (!enumerationType) {
                enumerationType = new EnumerationType(label: label, description: description)
                modelEnumTypes[label] = enumerationType
            }
            EnumerationValue enumerationValue = new EnumerationValue(key: key, value: value)
            addMetadataFromExtraColumns(enumerationValue, PluginSimpleExcelModule.ENUM_SHEET_COLUMNS, columnValues)
            enumerationType.addToEnumerationValues(enumerationValue)
        }
        return returnValues
    }

    //    static MODEl_SHEET_COLUMNS = ["DataClass Path", "Name", "Description", "Minimum Multiplicity", "Maximum Multiplicity", "DataType Name",
    //                                  "DataType Reference"]
    void addClassesAndElements(DataModel dataModel, Sheet dataModelSheet, Map<String, EnumerationType> enumerationTypes) {
        List<Map<String, String>> sheetValues = getSheetValues(PluginSimpleExcelModule.MODEL_SHEET_COLUMNS, dataModelSheet)
        Map<String, DataType> modelDataTypes = [:]
        sheetValues.each { row ->

            String dataClassPath = row["DataClass Path"]
            DataClass parentDataClass = getOrCreateClassFromPath(dataModel, dataClassPath)
            String name = row["Name"]
            String description = row["Description"]
            String minMult = row["Minimum Multiplicity"]
            String maxMult = row["Maximum Multiplicity"]
            String typeName = row["DataType Name"]
            String typeReference = row["DataType Reference"]
            MetadataAware createdElement = null
            if(!name || name=="") {
                // We're dealing with a data class
                createdElement = parentDataClass
                if(description != "") {
                    parentDataClass.description = description
                }
                if(minMult) {
                    parentDataClass.minMultiplicity = Integer.parseInt(minMult)
                }
                if(maxMult) {
                    if(maxMult == "*") {
                        parentDataClass.maxMultiplicity = -1
                    } else {
                        parentDataClass.maxMultiplicity = Integer.parseInt(maxMult)
                    }
                }
            } else {
                // We're dealing with a data element
                DataElement newDataElement = new DataElement(label: name, description: description)
                if(minMult) {
                    newDataElement.minMultiplicity = Integer.parseInt(minMult)
                }
                if(maxMult) {
                    if(maxMult == "*") {
                        newDataElement.maxMultiplicity = -1
                    } else {
                        newDataElement.maxMultiplicity = Integer.parseInt(maxMult)
                    }
                }
                DataType elementDataType
                if(enumerationTypes[typeName]) {
                    elementDataType = enumerationTypes[typeName]
                } else if( modelDataTypes[typeName]) {
                    elementDataType = modelDataTypes[typeName]
                }
                else {
                    if(typeReference) {
                        DataClass referenceDataClass = getOrCreateClassFromPath(dataModel, typeReference)
                        elementDataType = new ReferenceType(label: typeName)
                        referenceDataClass.addToReferenceTypes(elementDataType)
                    } else {
                        elementDataType = new PrimitiveType(label: typeName)
                    }
                    modelDataTypes[typeName] = elementDataType
                }
                newDataElement.dataType = elementDataType
                elementDataType.addToDataElements(newDataElement)
                parentDataClass.addToChildDataElements(newDataElement)
                createdElement = newDataElement
            }
            addMetadataFromExtraColumns(createdElement, PluginSimpleExcelModule.MODEL_SHEET_COLUMNS, row)

        }
        modelDataTypes.values().each {
            dataModel.addToDataTypes(it)
        }
        enumerationTypes.values().each {
            dataModel.addToDataTypes(it)
        }

    }

    DataClass getOrCreateClassFromPath(DataModel dataModel, String path) {
        List<String> pathComponents = path.split("\\|")
        DataClass parentDataClass = dataClassService.findOrCreateDataClassByPath(dataModel, pathComponents, "",
                                                                                                  null,
                                                                                                  null, null)
        return parentDataClass


    }


    String getCellValueAsString(Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }


}
