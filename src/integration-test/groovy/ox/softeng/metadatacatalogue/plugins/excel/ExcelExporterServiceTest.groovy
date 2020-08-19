package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelFileImporterParameters

import com.google.common.base.Strings
import org.junit.Test

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertFalse
import static org.junit.Assert.assertNotNull

/**
 * @since 01/03/2018
 */
@SuppressWarnings('SpellCheckingInspection')
class ExcelExporterServiceTest extends BaseDataModelImportExportExcelTest {

    @Test
    void testSimpleExport() {
        List<DataModel> dataModels = testExportViaImport('simpleImport.xlsx', 'simpleImport_export.xlsx')

        verifySimpleDataModel dataModels.first()
        verifySimpleDataModelContent dataModels.first()
    }

    @Test
    void testMultiModelExport() {
        List<DataModel> dataModels = testExportViaImport('multiDataModelImport.xlsx', 'multiDataModelImport_export.xlsx', 3)

        verifySimpleDataModel dataModels.find {it.label == 'test'}
        verifySimpleDataModelContent dataModels.find {it.label == 'test'}
    }

    @Test
    void testComplexMetadataModelExport() {
        List<DataModel> dataModels = testExportViaImport('simpleImportComplexMetadata.xlsx', 'simpleImportComplexMetadata_export.xlsx')

        DataModel dataModel = dataModels.first()
        verifySimpleDataModel dataModel, 4
        checkMetadata dataModel, 'ox.softeng.database|dialect', 'test'

        DataClass top = checkDataClass(dataModel, 'top', 3, 'tops description', 1, 1, [
            'extra info'          : 'some extra info',
            'ox.softeng.test|blob': 'hello'])
        checkDataElement top, 'info', 'info description', 'string', 1, 1, [
            'different info'    : 'info',
            'ox.softeng.xsd|min': '1',
            'ox.softeng.xsd|max': '20',
        ]
        checkDataElement top, 'another', null, 'int', 0, 1, ['extra info': 'some extra info', 'different info': 'info']
        checkDataElement top, 'more info', 'why not', 'string', 0

        DataClass child = checkDataClass(top, 'child', 4, 'child description', 0)
        checkDataElement child, 'info', 'childs info', 'string', 1, 1, [
            'extra info'          : 'some extra info',
            'ox.softeng.xsd|min'  : '0',
            'ox.softeng.xsd|max'  : '10',
            'ox.softeng.test|blob': 'wobble'
        ]
        checkDataElement child, 'does it work', 'I don\'t know', 'yesno', 0, -1, [
            'ox.softeng.xsd|choice': 'a',
        ]
        checkDataElement child, 'lazy', 'where the merging only happens in the id col', 'possibly', 1, 1, [
            'different info'       : 'info',
            'ox.softeng.xsd|choice': 'a',
        ]
        checkDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = checkDataClass(top, 'brother', 2)
        checkDataElement brother, 'info', 'brothers info', 'string', 0, 1, [
            'ox.softeng.xsd|min': '1',
            'ox.softeng.xsd|max': '255',]
        checkDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
    }

    private List<DataModel> testExportViaImport(String filename, String outFileName, int expectedCount = 1) throws IOException, ApiException {
        // Import model first
        ExcelFileImporterParameters params = createImportParameters(filename)
        List<DataModel> importedModels = importDomains(params, expectedCount)

        getLogger().debug('DataModel to export: {}', importedModels[0].getId())
        // Rather than use the one returned from the import, we want to check whats actually been saved into the DB

        Path outPath = Paths.get('build/tmp/', outFileName)

        if (importedModels.size() == 1) testExport(importedModels[0].getId(), outPath)
        else testExport(importedModels.id, outPath)

        getLogger().info('>>> Importing')
        params = createImportParameters(outPath)
        importDomains(params, expectedCount, false)
    }

    private void testExport(UUID dataModelId, Path outPath) throws IOException, ApiException {
        getLogger().info('>>> Exporting Single')
        ExcelDataModelExporterService exporterService = applicationContext.getBean(ExcelDataModelExporterService)
        ByteArrayOutputStream byteArrayOutputStream = exporterService.exportDomain(catalogueUser, dataModelId)
        assertNotNull('Should have an exported model', byteArrayOutputStream)

        String exported = byteArrayOutputStream.toString('ISO-8859-1')
        assertFalse('Should have an exported model string', Strings.isNullOrEmpty(exported))

        Files.write(outPath, exported.getBytes('ISO-8859-1'))
    }

    private void testExport(List<UUID> dataModelIds, Path outPath) throws IOException, ApiException {
        getLogger().info('>>> Exporting Multiple')
        ExcelDataModelExporterService exporterService = applicationContext.getBean(ExcelDataModelExporterService)
        ByteArrayOutputStream byteArrayOutputStream = exporterService.exportDomains(catalogueUser, dataModelIds)
        assertNotNull('Should have an exported model', byteArrayOutputStream)

        String exported = byteArrayOutputStream.toString('ISO-8859-1')
        assertFalse('Should have an exported model string', Strings.isNullOrEmpty(exported))

        Files.write(outPath, exported.getBytes('ISO-8859-1'))
    }
}
