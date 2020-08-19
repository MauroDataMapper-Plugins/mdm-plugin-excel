package ox.softeng.metadatacatalogue.plugins.simpleexcel

import ox.softeng.metadatacatalogue.core.spi.module.AbstractModule

/**
 * @since 17/08/2017
 */
@SuppressWarnings('SpellCheckingInspection')
class PluginSimpleExcelModule extends AbstractModule {

    @Override
    String getName() {
        'Plugin - Simple Excel'
    }

    //    @Override
    Closure doWithSpring() {
        {->
            simpleExcelImporterService(SimpleExcelDataModelImporterService)
            simpleExcelExporterService(SimpleExcelDataModelExporterService)

        }
    }

    static DATAMODEL_SHEET_COLUMNS = ["Name", "Description", "Author", "Organisation", "Sheet Key", "Type"]
    static ENUM_SHEET_COLUMNS = ["DataModel Name", "Name", "Description", "Key", "Value"]
    static MODEL_SHEET_COLUMNS = ["DataClass Path", "Name", "Description", "Minimum Multiplicity", "Maximum Multiplicity", "DataType Name",
                                  "DataType Reference"]

}
