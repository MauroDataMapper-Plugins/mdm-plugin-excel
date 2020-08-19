package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.catalogue.CatalogueItem
import ox.softeng.metadatacatalogue.core.facet.Metadata
import ox.softeng.metadatacatalogue.core.spi.importer.ImporterPlugin
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.ImporterPluginParameters
import ox.softeng.metadatacatalogue.plugins.test.BaseImportPluginTest

import org.grails.datastore.gorm.GormEntity

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.fail

/**
 * @since 09/03/2018
 */
abstract class BaseImportExportExcelTest<D extends GormEntity, P extends ImporterPluginParameters, T extends ImporterPlugin<D, P>>
    extends BaseImportPluginTest<D, P, T> {

    protected void checkMetadata(CatalogueItem catalogueItem, String key, String value) {

        String[] parts = key.split(/\|/)
        String ns = importerInstance.namespace
        String k = parts[0]
        if (parts.size() != 1) {
            ns = parts[0]
            k = parts[1]
        }

        Metadata md = catalogueItem.findMetadataByNamespaceAndKey(ns, k)
        assertNotNull("${catalogueItem.label} Metadata ${key} exists", md)
        assertEquals("${catalogueItem.label} Metadata ${key} value", value, md.value)
    }

    protected P createImportParameters(String filename) throws IOException {
        Path p = Paths.get('src/integration-test/resources/' + filename)
        if (!Files.exists(p)) {
            fail("File ${filename} cannot be found")
        }
        createImportParameters(p)
    }

    abstract protected P createImportParameters(Path path) throws IOException
}
