package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.api.exception.ApiBadRequestException
import ox.softeng.metadatacatalogue.core.api.exception.ApiUnauthorizedException
import ox.softeng.metadatacatalogue.core.catalogue.CatalogueItem
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClassService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElement
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElementService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.DataType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.PrimitiveTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.ReferenceTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.spi.importer.DataModelImporterPlugin
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.FileParameter
import ox.softeng.metadatacatalogue.core.type.catalogue.DataModelType
import ox.softeng.metadatacatalogue.core.user.CatalogueUser
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelFileImporterParameters
import ox.softeng.metadatacatalogue.plugins.excel.row.catalogue.ContentDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.catalogue.DataModelDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.column.MetadataColumn
import ox.softeng.metadatacatalogue.plugins.excel.util.WorkbookHandler

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.beans.factory.annotation.Autowired

import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.CONTENT_HEADER_ROWS
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.CONTENT_ID_COLUMN
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATAMODELS_HEADER_ROWS
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATAMODELS_ID_COLUMN
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATAMODELS_SHEET_NAME

/**
 * @since 01/03/2018
 */
class ExcelDataModelImporterService implements DataModelImporterPlugin<ExcelFileImporterParameters>, WorkbookHandler {

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
    DataModel importDataModel(CatalogueUser currentUser, ExcelFileImporterParameters params) {
        importDataModels(currentUser, params)?.first()
    }

    @Override
    List<DataModel> importDataModels(CatalogueUser currentUser, ExcelFileImporterParameters importerParameters) {
        if (!currentUser) throw new ApiUnauthorizedException('EISP01', 'User must be logged in to import model')
        if (importerParameters.importFile.fileContents.size() == 0) throw new ApiBadRequestException('EIS02', 'Cannot import empty file')
        logger.info('Importing {} as {}', importerParameters.importFile.fileName, currentUser.emailAddress)

        FileParameter importFile = importerParameters.importFile

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromInputStream(importFile.fileName, importFile.inputStream)

            List<DataModelDataRow> dataModelDataRows = loadDataRows(workbook, DataModelDataRow, importFile.fileName,
                                                                    DATAMODELS_SHEET_NAME, DATAMODELS_HEADER_ROWS,
                                                                    DATAMODELS_ID_COLUMN)
            if (!dataModelDataRows) return []

            List<DataModel> dataModels = loadDataModels(dataModelDataRows, currentUser, workbook, importFile.fileName)

            return dataModels.findAll {it.dataClasses}
        } finally {
            closeWorkbook workbook
        }
    }

    List<DataModel> loadDataModels(List<DataModelDataRow> dataRows, CatalogueUser catalogueUser, Workbook workbook, String filename) {
        List<DataModel> dataModels = []

        dataRows.each {dataRow ->
            logger.debug('Creating DataModel {}', dataRow.name)
            DataModel dataModel = dataModelService.createDataModel(catalogueUser, dataRow.name, dataRow.description, dataRow.author,
                                                                   dataRow.organisation, DataModelType.findForLabel(dataRow.type))

            addMetadataToCatalogueItem dataModel, dataRow.metadata, catalogueUser

            if (!workbook.getSheet(dataRow.sheetKey)) {
                logger.warn('DataModel {} with key {} will not be imported as it has no corresponding sheet', dataRow.name, dataRow.sheetKey)
                return
            }

            List<ContentDataRow> contentDataRows = loadDataRows(workbook, ContentDataRow, filename, dataRow.sheetKey,
                                                                CONTENT_HEADER_ROWS, CONTENT_ID_COLUMN)

            if (!contentDataRows) return

            dataModels += loadSheetDataRowsIntoDataModel(contentDataRows, dataModel, catalogueUser)
        }

        dataModels
    }

    DataModel loadSheetDataRowsIntoDataModel(List<ContentDataRow> dataRows, DataModel dataModel, CatalogueUser catalogueUser) {

        // Load DataClasses
        loadDataClassRows(dataRows, dataModel, catalogueUser)

        // Load the rest of the rows as DataElements
        loadDataElementRows(dataRows.findAll {it.dataElementName}, dataModel, catalogueUser)

        dataModel
    }

    void loadDataClassRows(List<ContentDataRow> dataRows, DataModel dataModel, CatalogueUser catalogueUser) {
        logger.info('Loading DataClass rows for DataModel {}', dataModel.label)

        // Where not DataElement name then the row is DataClass specific data, otherwise we just need to make sure all the DataClasses exist
        // Sort the rows so we generate them in path descending order, this will ensure we create any specific data before we create children
        dataRows.groupBy {it.dataClassPath}.sort().each {path, dataClassRows ->

            // Find the row with no DataElement name, if it doesn't exist then just grab the first row
            ContentDataRow dataRow = dataClassRows.find {!it.dataElementName} ?: dataClassRows.first()

            String description = dataRow.dataElementName ? null : dataRow.description
            Integer min = dataRow.dataElementName ? 1 : dataRow.minMultiplicity
            Integer max = dataRow.dataElementName ? 1 : dataRow.maxMultiplicity

            logger.debug('Creating DataClass [{}]', dataRow.dataClassPath)
            DataClass dataClass = dataClassService.findOrCreateDataClassByPath(dataModel, dataRow.dataClassPathList, description, catalogueUser,
                                                                               min, max)
            if (!dataRow.dataElementName) {
                addMetadataToCatalogueItem dataClass, dataRow.metadata, catalogueUser
            }
        }
    }

    void loadDataElementRows(List<ContentDataRow> dataRows, DataModel dataModel, CatalogueUser catalogueUser) {
        dataRows.each {dataRow ->

            DataClass parentDataClass = dataClassService.findDataClassByPath(dataModel, dataRow.dataClassPathList)
            DataType dataType = null

            if (dataRow.mergedContentRows) {
                logger.debug('DataElement is an EnumerationType')

                dataType = enumerationTypeService.findOrCreateDataTypeForDataModel(dataModel, dataRow.dataTypeName, dataRow.dataTypeDescription,
                                                                                   catalogueUser)

                dataRow.mergedContentRows.each {mcr ->
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

            logger.debug('Adding DataElement [{}] to DataClass [{}] with DataType [{}]', dataRow.dataElementName, parentDataClass?.label,
                         dataType?.label)
            DataElement dataElement = dataElementService.findOrCreateDataElementForDataClass(
                parentDataClass, dataRow.dataElementName, dataRow.description, catalogueUser, dataType, dataRow.minMultiplicity,
                dataRow.maxMultiplicity
            )

            addMetadataToCatalogueItem dataElement, dataRow.metadata, catalogueUser
        }

        dataModel
    }

    void addMetadataToCatalogueItem(CatalogueItem catalogueItem, List<MetadataColumn> metadata, CatalogueUser catalogueUser) {
        metadata.each {md ->
            logger.trace('Adding Metadata {}', md.key)
            catalogueItem.addToMetadata(md.namespace ?: namespace, md.key, md.value, catalogueUser)
        }
    }
}
