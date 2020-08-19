package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlow
import ox.softeng.metadatacatalogue.core.catalogue.dataflow.DataFlowComponent
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.FileParameter
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelDataFlowImporterParameters
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelFileImporterParameters

import com.google.common.base.Strings
import org.hibernate.SessionFactory
import org.junit.Test

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertFalse
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.assertTrue
import static org.junit.Assert.fail

/**
 * @since 09/03/2018
 */
class ExcelDataFlowExporterServiceTest extends BaseImportExportExcelTest<DataFlow, ExcelDataFlowImporterParameters, ExcelDataFlowImporterService> {


    @Test
    void testSimpleExport() {
        importTestDataModels()

        List<DataFlow> dataFlows = testExportViaImport('Simple_DataFlow_Import_File.xlsx', 'Simple_DataFlow_Import_File_Export.xlsx')

        verifySimpleDataFlow dataFlows.first()
    }

    @Test
    void performMultiDataFlowExport() {
        importTestDataModels()

        List<DataFlow> dataFlows = testExportViaImport('Multiple_DataFlow_Import_File.xlsx', 'Multiple_DataFlow_Import_File_Export.xlsx', 2)

        DataFlow dataFlow = dataFlows.find {it.label == 'Sample DataFlow'}

        verifySimpleDataFlow(dataFlow)

        dataFlow = dataFlows.find {it.label == 'Second DataFlow'}

        assertNotNull('Second model must exist', dataFlow)

        assertEquals('DataFlow name', 'Second DataFlow', dataFlow.label)
        assertEquals('DataFlow description', 'This is a second dataflow between the another model and complex model',
                     dataFlow.description)

        assertEquals('DataFlow source', 'complex.xsd', dataFlow.source.label)
        assertEquals('DataFlow target', 'Another Model', dataFlow.target.label)

        assertEquals("Number of DataFlow Components", 9, dataFlow.dataFlowComponents.size())
    }

    protected void verifySimpleDataFlow(DataFlow dataFlow) {
        assertNotNull('Simple model must exist', dataFlow)
        assertEquals('DataFlow name', 'Sample DataFlow', dataFlow.label)
        assertEquals('DataFlow description', 'This is a sample dataflow which shows all the possible ways we can show data',
                     dataFlow.description)

        assertEquals('DataFlow source', 'test', dataFlow.source.label)
        assertEquals('DataFlow target', 'Another Model', dataFlow.target.label)

        assertEquals("Number of DataFlow Components", 8, dataFlow.dataFlowComponents.size())

        Collection<DataFlowComponent> dfcs = dataFlow.dataFlowComponents.findAll {it.label == 'merge'}

        assertEquals("Number of merge transforms", 2, dfcs.size())

        assertTrue("Merge with 3 source & 1 target", dfcs.any {
            it.sourceElements.size() == 3 && it.targetElements.size() == 1 &&
            it.description == 'merge all element info to one location, we will be keeping the elements but we want all the info together as ' +
            'well'
        })
        assertTrue("Merge with 2 source & 1 target", dfcs.any {
            def md = it.metadata ?: []
            it.sourceElements.size() == 2 && it.targetElements.size() == 1 &&
            it.description == 'merge 2 info elements into one' &&
            md.size() == 1 &&
            md.first().key == 'merge separator' &&
            md.first().value == ','
        })

        dfcs = dataFlow.dataFlowComponents.findAll {it.label == 'copy'}

        assertEquals("Number of copy transforms", 3, dfcs.size())

        assertTrue("Copy another", dfcs.any {it.sourceElements.first().label == 'another' && it.targetElements.first().label == 'another'})
        assertTrue("Copy info child", dfcs.any {
            it.sourceElements.first().label == 'info' &&
            it.targetElements.first().label == 'info' &&
            it.sourceElements.first().dataClass.label == 'child' &&
            it.targetElements.first().dataClass.label == 'child'
        })
        assertTrue("Copy info brother", dfcs.any {
            it.sourceElements.first().label == 'info' &&
            it.targetElements.first().label == 'info' &&
            it.sourceElements.first().dataClass.label == 'brother' &&
            it.targetElements.first().dataClass.label == 'brother'
        })
        assertTrue("Copy description", dfcs.every {it.description == 'copy with no changes'})

        DataFlowComponent dfc = dataFlow.dataFlowComponents.find {it.label == 'required field'}

        assertNotNull("Required field should exist", dfc)
        assertEquals('required field description', 'if empty then set to the default', dfc.description)
        assertEquals('required field source', 'does it work', dfc.sourceElements.first().label)
        assertEquals('required field target', 'does it work', dfc.targetElements.first().label)
        checkMetadata(dfc, 'default', 'y')

        dfc = dataFlow.dataFlowComponents.find {it.label == 'clean'}

        assertNotNull("clean should exist", dfc)
        assertEquals('clean description', 'this field contains dangerous data which needs to be cleaned', dfc.description)
        assertEquals('clean source', 'wrong order', dfc.sourceElements.first().label)
        assertEquals('clean target', 'wrong order', dfc.targetElements.first().label)
        checkMetadata(dfc, 'cleaning technique', 'all identifiable data')

        dfc = dataFlow.dataFlowComponents.find {it.label == 'duplicate'}

        assertNotNull("duplicate should exist", dfc)
        assertEquals('duplicate description', 'we need to duplicate this data into another field', dfc.description)
        assertEquals('duplicate source', 'sibling', dfc.sourceElements.first().label)
        assertEquals('duplicate targets size', 2, dfc.targetElements.size())
        assertTrue('duplicate target 1', dfc.targetElements.any {it.label == 'sibling' && it.dataClass.label == 'brother'})
        assertTrue('duplicate target 2', dfc.targetElements.any {it.label == 'twin_sibling' && it.dataClass.label == 'brother'})
    }

    @Override
    protected ExcelDataFlowImporterParameters createImportParameters(Path path) throws IOException {
        ExcelDataFlowImporterParameters params = new ExcelDataFlowImporterParameters()
        FileParameter file = new FileParameter(path.toString(), 'application/vnd.ms-excel', Files.readAllBytes(path))

        params.setImportFile(file)
        params
    }

    protected void importTestDataModels() {
        try {
            ExcelDataModelImporterService excelImporterService = getBean(ExcelDataModelImporterService)

            Path p = Paths.get('src/integration-test/resources/multiDataModelImport.xlsx')
            if (!Files.exists(p)) {
                fail('File multiDataModelImport.xlsx cannot be found')
            }
            ExcelFileImporterParameters params = new ExcelFileImporterParameters(finalised: false)
            FileParameter file = new FileParameter(p.toString(), 'application/vnd.ms-excel', Files.readAllBytes(p))

            params.setImportFile(file)
            List<DataModel> dataModels = excelImporterService.importDataModels(catalogueUser, params)
            dataModels*.save(flush: true)
        } finally {
            SessionFactory sessionFactory = getBean(SessionFactory)
            sessionFactory.getCurrentSession().clear()
        }
    }

    private List<DataFlow> testExportViaImport(String filename, String outFileName, int expectedCount = 1) throws IOException, ApiException {
        // Import model first
        ExcelDataFlowImporterParameters params = createImportParameters(filename)
        List<DataFlow> importedModels = importDomains(params, expectedCount)

        getLogger().debug('DataFlow to export: {}', importedModels[0].getId())
        // Rather than use the one returned from the import, we want to check whats actually been saved into the DB

        Path outPath = Paths.get('build/tmp/', outFileName)

        if (importedModels.size() == 1) testExport(importedModels[0].getId(), outPath)
        else testExport(importedModels.id, outPath)

        getLogger().info('>>> Importing')
        params = createImportParameters(outPath)
        importDomains(params, expectedCount, false)
    }

    private void testExport(UUID dataFlowId, Path outPath) throws IOException, ApiException {
        getLogger().info('>>> Exporting Single')
        ExcelDataFlowExporterService exporterService = applicationContext.getBean(ExcelDataFlowExporterService)
        ByteArrayOutputStream byteArrayOutputStream = exporterService.exportDomain(catalogueUser, dataFlowId)
        assertNotNull('Should have an exported model', byteArrayOutputStream)

        String exported = byteArrayOutputStream.toString('ISO-8859-1')
        assertFalse('Should have an exported model string', Strings.isNullOrEmpty(exported))

        Files.write(outPath, exported.getBytes('ISO-8859-1'))
    }

    private void testExport(List<UUID> dataFlowIds, Path outPath) throws IOException, ApiException {
        getLogger().info('>>> Exporting Multiple')
        ExcelDataFlowExporterService exporterService = applicationContext.getBean(ExcelDataFlowExporterService)
        ByteArrayOutputStream byteArrayOutputStream = exporterService.exportDomains(catalogueUser, dataFlowIds)
        assertNotNull('Should have an exported model', byteArrayOutputStream)

        String exported = byteArrayOutputStream.toString('ISO-8859-1')
        assertFalse('Should have an exported model string', Strings.isNullOrEmpty(exported))

        Files.write(outPath, exported.getBytes('ISO-8859-1'))
    }

    @Override
    DataFlow saveDomain(DataFlow domain) {
        domain.save(validate: false, flush: true)
    }
}
