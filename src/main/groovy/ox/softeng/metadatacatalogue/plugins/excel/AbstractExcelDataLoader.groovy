package ox.softeng.metadatacatalogue.plugins.excel

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.core.GenericTypeResolver
import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClassService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElementService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.PrimitiveTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.spi.dataloader.DataModelDataLoaderPlugin
import ox.softeng.metadatacatalogue.core.user.CatalogueUser
import ox.softeng.metadatacatalogue.plugins.excel.row.StandardDataRow
import ox.softeng.metadatacatalogue.plugins.excel.util.WorkbookHandler

/**
 * @since 15/02/2018
 */
abstract class AbstractExcelDataLoader<K extends StandardDataRow> implements DataModelDataLoaderPlugin, WorkbookHandler {

    @Autowired
    DataModelService dataModelService

    @Autowired
    PrimitiveTypeService primitiveTypeService

    @Autowired
    DataClassService dataClassService

    @Autowired
    EnumerationTypeService enumerationTypeService

    @Autowired
    DataElementService dataElementService

    abstract String getExcelFilename()

    abstract List<String> getExcelSheetNames()

    abstract String getDataModelName()

    abstract int getNumberOfHeaderRows()

    abstract int getIdCellNumber()

    abstract DataModel loadSheetDataRowsIntoDataModel(String sheetName, List<K> dataRows, DataModel dataModel, CatalogueUser catalogueUser)

    @Override
    List<DataModel> importData(CatalogueUser catalogueUser) throws ApiException {
        logger.info('Importing {} Model', name)

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromFilename(getExcelFilename())

            DataModel dataModel = dataModelService.createDataModel(catalogueUser, getDataModelName(), getDescription(), getAuthor(),
                                                                   getOrganisation())
            Class<K> dataRowClass = (Class<K>) GenericTypeResolver.resolveTypeArgument(getClass(), AbstractExcelDataLoader)
            // Load each sheet
            excelSheetNames.each {sheetName ->
                logger.debug('Loading sheet {} from file {}', sheetName, getExcelFilename())
                List<K> dataRows = loadDataRowsFromSheet(workbook, dataRowClass, sheetName)
                if (!dataRows) return
                loadSheetDataRowsIntoDataModel(sheetName, dataRows, dataModel, catalogueUser)
            }

            dataModel.dataClasses ? [dataModel] : []

        } finally {
            closeWorkbook workbook
        }
    }

    List<K> loadDataRowsFromSheet(Workbook workbook, Class<K> dataRowClass, String sheetName) {
        loadDataRows(workbook, dataRowClass, getExcelFilename(), sheetName, getNumberOfHeaderRows(),
                     getIdCellNumber())
    }
}
