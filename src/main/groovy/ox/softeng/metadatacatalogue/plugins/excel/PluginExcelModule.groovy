package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.spi.module.AbstractModule

import java.awt.Color

/**
 * @since 17/08/2017
 */
@SuppressWarnings('SpellCheckingInspection')
class PluginExcelModule extends AbstractModule {

    public static final String HEADER_COLUMN_COLOUR = '#5B9BD5'
    public static final Color ALTERNATING_COLUMN_COLOUR = Color.decode('#7BABF5')
    public static final Color BORDER_COLOUR = Color.decode('#FFFFFF')
    public static final Double CELL_COLOUR_TINT = 0.6d
    public static final Double BORDER_COLOUR_TINT = -0.35d

    public static final String DATAMODELS_SHEET_NAME = 'DataModels'
    public static final Integer DATAMODELS_HEADER_ROWS = 2
    public static final Integer DATAMODELS_ID_COLUMN = 1

    public static final Integer CONTENT_HEADER_ROWS = 2
    public static final Integer CONTENT_ID_COLUMN = 0

    public static final String DATAFLOWS_SHEET_NAME = 'DataFlows'
    public static final Integer DATAFLOWS_HEADER_ROWS = 2
    public static final Integer DATAFLOWS_ID_COLUMN = 1

    public static final Integer DATAFLOWCOMPONENT_HEADER_ROWS = 2
    public static final Integer DATAFLOWCOMPONENT_ID_COLUMN = 2

    public static final String DATAMODEL_TEMPLATE_FILENAME = 'Template_DataModel_Import_File.xlsx'
    public static final String DATAFLOW_TEMPLATE_FILENAME = 'Template_DataFlow_Import_File.xlsx'

    public static final String DATACLASS_PATH_SPLIT_REGEX = ~/\|/

    @Override
    String getName() {
        'Plugin - Excel'
    }

    //    @Override
    Closure doWithSpring() {
        {->
            excelImporterService(ExcelDataModelImporterService)
            excelExporterService(ExcelDataModelExporterService)
            excelDataFlowImporterService(ExcelDataFlowImporterService)
            excelDataFlowExporterService(ExcelDataFlowExporterService)
        }
    }

}
