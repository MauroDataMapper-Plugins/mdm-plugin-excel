package ox.softeng.metadatacatalogue.plugins.simpleexcel.parameter

import ox.softeng.metadatacatalogue.core.spi.importer.parameter.DataModelFileImporterPluginParameters

/**
 * @since 01/03/2018
 */
class SimpleExcelFileImporterParameters extends DataModelFileImporterPluginParameters {

    // So we don't ask for the DataModel name
    String dataModelName
}
