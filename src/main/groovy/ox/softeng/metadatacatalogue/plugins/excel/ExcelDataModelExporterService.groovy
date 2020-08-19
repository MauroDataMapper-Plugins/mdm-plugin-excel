package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.catalogue.CatalogueItem
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClassService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElementService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.spi.exporter.DataModelExporterPlugin
import ox.softeng.metadatacatalogue.core.user.CatalogueUser
import ox.softeng.metadatacatalogue.plugins.excel.row.catalogue.ContentDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.catalogue.DataModelDataRow
import ox.softeng.metadatacatalogue.plugins.excel.row.catalogue.EnumerationValueDataRow
import ox.softeng.metadatacatalogue.plugins.excel.util.WorkbookExporter

import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.beans.factory.annotation.Autowired

import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.ALTERNATING_COLUMN_COLOUR
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.CELL_COLOUR_TINT
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.CONTENT_HEADER_ROWS
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATAMODELS_HEADER_ROWS
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATAMODELS_SHEET_NAME
import static ox.softeng.metadatacatalogue.plugins.excel.PluginExcelModule.DATAMODEL_TEMPLATE_FILENAME

/**
 * @since 01/03/2018
 */
class ExcelDataModelExporterService implements DataModelExporterPlugin, WorkbookExporter {

    @Autowired
    DataModelService dataModelService

    @Autowired
    DataElementService dataElementService

    @Autowired
    DataClassService dataClassService

    @Override
    String getDisplayName() {
        'Excel (XLSX) Exporter'
    }

    @Override
    String getVersion() {
        '1.0.0'
    }

    @Override
    Boolean canExportMultipleDomains() {
        true
    }

    @Override
    ByteArrayOutputStream exportDataModel(CatalogueUser currentUser, DataModel dataModel) throws ApiException {
        exportDataModels(currentUser, [dataModel])
    }

    @Override
    ByteArrayOutputStream exportDataModels(CatalogueUser currentUser, List<DataModel> dataModels) throws ApiException {

        logger.info('Exporting DataModels to Excel')
        XSSFWorkbook workbook = null
        try {
            workbook = loadWorkbookFromFilename(DATAMODEL_TEMPLATE_FILENAME) as XSSFWorkbook

            workbook = loadDataModelDataRowsIntoWorkbook(dataModels.collect {new DataModelDataRow(it)}, workbook)

            if (workbook) {
                ByteArrayOutputStream os = new ByteArrayOutputStream()
                workbook.write(os)
                logger.info('DataModels exported')
                return os
            }
        } finally {
            closeWorkbook(workbook)
        }
        null
    }

    XSSFWorkbook loadDataModelDataRowsIntoWorkbook(List<DataModelDataRow> dataModelDataRows, XSSFWorkbook workbook) {
        XSSFColor cellColour = new XSSFColor(ALTERNATING_COLUMN_COLOUR)
        cellColour.setTint(CELL_COLOUR_TINT)

        XSSFCellStyle defaultStyle = createDefaultCellStyle(workbook as XSSFWorkbook)

        XSSFCellStyle colouredStyle = workbook.createCellStyle() as XSSFCellStyle
        colouredStyle.cloneStyleFrom(defaultStyle)
        colouredStyle.setFillForegroundColor(cellColour)
        colouredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)

        Sheet dataModelSheet = workbook.getSheet(DATAMODELS_SHEET_NAME)

        addMetadataHeadersToSheet(dataModelSheet, dataModelDataRows)

        dataModelDataRows.each {dataModelDataRow ->

            logger.debug('Loading DataModel {} into workbook', dataModelDataRow.name)

            // We need to allow for potential sheet name already exists
            Sheet contentSheet = createComponentSheetFromTemplate(workbook, dataModelDataRow.sheetKey)
            dataModelDataRow.sheetKey = contentSheet.sheetName

            logger.debug('Adding DataModel row')
            dataModelDataRow.buildRow(dataModelSheet.createRow(dataModelSheet.lastRowNum + 1))

            logger.debug('Configuring {} sheet style', dataModelSheet.sheetName)
            configureSheetStyle(dataModelSheet, dataModelDataRow.firstMetadataColumn, DATAMODELS_HEADER_ROWS, defaultStyle, colouredStyle)

            logger.debug('Adding DataModel content to sheet {}', dataModelDataRow.sheetKey)
            List<ContentDataRow> contentDataRows = createContentDataRows(dataModelDataRow.dataModel.childDataClasses.sort {it.label})

            if (contentDataRows) {
                loadDataRowsIntoSheet(contentSheet, contentDataRows)

                logger.debug('Configuring {} sheet style', contentSheet.sheetName)
                configureSheetStyle(contentSheet, contentDataRows.first().firstMetadataColumn, CONTENT_HEADER_ROWS, defaultStyle, colouredStyle)
            } else {
                logger.warn('No content rows to add for DataModel')
            }

        }
        removeTemplateSheet(workbook)
    }

    List<ContentDataRow> createContentDataRows(List<CatalogueItem> catalogueItems, List<ContentDataRow> contentDataRows = []) {
        logger.debug('Creating content DataRows')
        catalogueItems.each {item ->
            logger.trace('Creating content {}:{}', item.domainType, item.label)
            ContentDataRow dataRow = new ContentDataRow(item)
            contentDataRows += dataRow
            if (item.instanceOf(DataClass)) {
                contentDataRows = createContentDataRows(dataElementService.findAllByDataClassIdJoinDataType(item.id), contentDataRows)
                contentDataRows = createContentDataRows(dataClassService.findAllByParentDataClassId(item.id, [sort: 'label']), contentDataRows)
            }
        }
        contentDataRows
    }

    @Override
    Integer configureExtraRowStyle(Row row, int metadataColumnIndex, int lastColumnIndex, XSSFCellStyle defaultStyle,
                                   XSSFCellStyle colouredStyle, XSSFCellStyle mainCellStyle, XSSFCellStyle borderCellStyle,
                                   XSSFCellStyle metadataBorderStyle, XSSFCellStyle metadataColouredBorderStyle, boolean colourRow) {
        CellRangeAddress mergedRegion = getMergeRegion(row.getCell(0))
        if (mergedRegion) {
            boolean internalColourRow = colourRow
            for (int m = mergedRegion.firstRow; m <= mergedRegion.lastRow; m++) {
                Row internalRow = row.sheet.getRow(m)
                setRowStyle(internalRow, mainCellStyle, borderCellStyle, 0, EnumerationValueDataRow.KEY_COL_INDEX, metadataColumnIndex)
                setRowStyle(internalRow, internalColourRow ? colouredStyle : defaultStyle,
                            internalColourRow ? metadataColouredBorderStyle : metadataBorderStyle,
                            EnumerationValueDataRow.KEY_COL_INDEX, EnumerationValueDataRow.VALUE_COL_INDEX + 1, metadataColumnIndex)
                setRowStyle(internalRow, mainCellStyle, borderCellStyle, EnumerationValueDataRow.VALUE_COL_INDEX + 1, lastColumnIndex,
                            metadataColumnIndex)

                internalColourRow = !internalColourRow
            }
            return mergedRegion.lastRow
        }
        row.rowNum
    }
}